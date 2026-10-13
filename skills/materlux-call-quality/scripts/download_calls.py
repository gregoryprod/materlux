"""Daily read-only Bitrix24 call collector. Google Drive uploads are handled by the chat connector."""
import argparse
import hashlib
import json
import re
import sys
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlsplit
from read_check import Client, SafeError
from probe_calls import read_call

import os
ROOT = Path(os.environ['MATERLUX_OUTPUT_ROOT']).resolve()
PORTAL_HOST = urlsplit(os.environ['MATERLUX_WEBHOOK']).hostname
MSK = timezone(timedelta(hours=3))


def save_json(path, value):
    temp = path.with_suffix(path.suffix + '.tmp')
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    temp.replace(path)


def normalized(value):
    return ' '.join(str(value).lower().replace('ё', 'е').split())


def crm_link(type_id, entity_id):
    entity = {'1': 'lead', '2': 'deal', '3': 'contact', '4': 'company'}.get(str(type_id))
    return f'https://{PORTAL_HOST}/crm/{entity}/details/{int(entity_id)}/' if entity and str(entity_id).isdigit() and int(entity_id) > 0 else ''


def validate_audio(path):
    size = path.stat().st_size
    with path.open('rb') as stream:
        head = stream.read(16)
    valid = head.startswith(b'ID3') or (len(head) > 1 and head[0] == 255 and (head[1] & 224) == 224) or head.startswith(b'OggS') or head.startswith(b'fLaC') or (head.startswith(b'RIFF') and head[8:12] == b'WAVE') or head[4:8] == b'ftyp'
    if not valid or not size:
        raise SafeError('NOT_RECOGNIZED_AUDIO')
    return hashlib.sha256(path.read_bytes()).hexdigest()


class AudioRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parsed = urlsplit(newurl)
        host = (parsed.hostname or '').lower()
        if parsed.scheme != 'https' or parsed.username or parsed.password or not (
                host == PORTAL_HOST or host.endswith('.bitrix24.ru') or
                host.endswith('.bitrix24.com') or host.endswith('.b24cdn.com')):
            raise SafeError('AUDIO_REDIRECT_BLOCKED')
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def run(day, limit, allow_partial=False):
    snapshot = json.loads((ROOT / 'Сотрудники_источник.json').read_text(encoding='utf-8'))
    # Every scheduled run must refresh this snapshot through Google Sheets first.
    read_at = datetime.fromisoformat(snapshot['read_at'].replace('Z', '+00:00'))
    if (datetime.now(timezone.utc) - read_at).total_seconds() > 7200:
        raise SafeError('EMPLOYEE_SOURCE_STALE_REFRESH_GOOGLE_SHEETS')
    employees = [dict(direction=r[0], name=r[1], status=r[3] if len(r)>3 else '')
                 for r in snapshot['values'][1:] if len(r)>1 and str(r[1]).strip()]
    start = datetime.fromisoformat(day).replace(tzinfo=MSK)
    end = start + timedelta(days=1)
    out = ROOT / day
    out.mkdir(parents=True, exist_ok=True)
    manifest_path = out / 'Журнал выгрузки.json'
    manifest = json.loads(manifest_path.read_text(encoding='utf-8')) if manifest_path.exists() else {
        'date_msk': day, 'source': snapshot['source_url'], 'calls': {}, 'google_folder_id': None}
    client = Client()
    profile = client.call('profile', {})['result']
    if not employees:
        raise SafeError('EMPLOYEE_LIST_EMPTY')
    surnames = sorted({e['name'].split()[-1] for e in employees})
    users = []
    cursor = 0
    while True:
        data = read_call(client, 'user.get', {'FILTER': {'LAST_NAME': surnames},
             'select': ['ID', 'NAME', 'LAST_NAME', 'ACTIVE'], 'start': cursor})
        users.extend(data['result'])
        if 'next' not in data:
            break
        cursor = int(data['next'])
    mapping, problems = {}, []
    for employee in employees:
        matches = [u for u in users if normalized(str(u.get('NAME', '')) + ' ' + str(u.get('LAST_NAME', ''))) == normalized(employee['name'])]
        if not matches:
            fallback = read_call(client, 'user.get', {'FILTER': {'NAME_SEARCH': employee['name'].split()[-1]},
                                  'select': ['ID', 'NAME', 'LAST_NAME', 'ACTIVE']})
            candidates = fallback['result']
            matches = [u for u in candidates if sorted(normalized(str(u.get('NAME', '')) + ' ' + str(u.get('LAST_NAME', ''))).split()) == sorted(normalized(employee['name']).split())]
            if not matches:
                users.extend(candidates)
        if len(matches) != 1:
            problems.append({'name': employee['name'], 'matching_users': len(matches)})
        else:
            mapping[str(matches[0]['ID'])] = employee
    manifest['employee_mapping'] = mapping
    manifest['employee_mapping_problems'] = problems
    manifest['connection_user_id'] = profile['ID']
    save_json(manifest_path, manifest)
    if problems:
        print(json.dumps({'mapping_problems': problems, 'returned_users': [
            {'ID': u.get('ID'), 'NAME': u.get('NAME'), 'LAST_NAME': u.get('LAST_NAME')} for u in users]}, ensure_ascii=False), flush=True)
        if not allow_partial:
            raise SafeError('EMPLOYEE_MAPPING_INCOMPLETE')
    rows, cursor = [], 0
    expected_total = None
    while True:
        data = read_call(client, 'voximplant.statistic.get', {'FILTER': {
             '>=CALL_START_DATE': start.isoformat(), '<CALL_START_DATE': end.isoformat(),
             'PORTAL_USER_ID': list(mapping), 'CALL_FAILED_CODE': '200',
             '>=CALL_DURATION': 60, 'CALL_TYPE': [1, 2]},
             'SORT': 'ID', 'ORDER': 'ASC', 'start': cursor})
        rows.extend(data['result'])
        expected_total = data.get('total', expected_total)
        if 'next' not in data:
            break
        next_cursor = int(data['next'])
        if next_cursor <= cursor:
            raise SafeError('PAGINATION_NOT_ADVANCING')
        cursor = next_cursor
    if len({str(r['ID']) for r in rows}) != len(rows) or (expected_total is not None and len(rows) != int(expected_total)):
        raise SafeError('CALL_COUNT_MISMATCH')
    manifest['eligible_calls'] = len(rows)
    manifest['last_run_at'] = datetime.now(MSK).isoformat()
    print(json.dumps({'date_msk': day, 'mapped_employees': len(mapping), 'eligible_calls': len(rows)}, ensure_ascii=False), flush=True)
    processed = 0
    for row in rows:
        call_id = str(row['ID'])
        employee = mapping.get(str(row['PORTAL_USER_ID']))
        when = datetime.fromisoformat(row['CALL_START_DATE']).astimezone(MSK)
        if not employee or not start <= when < end or int(row['CALL_DURATION']) < 60 or str(row['CALL_TYPE']) not in {'1', '2'} or str(row['CALL_FAILED_CODE']) != '200':
            raise SafeError('FILTER_RESULT_MISMATCH')
        item = manifest['calls'].setdefault(call_id, {})
        item.update({'call_id': call_id, 'date': day, 'time_msk': when.strftime('%H:%M:%S'),
            'employee': employee['name'], 'employee_id': str(row['PORTAL_USER_ID']),
            'direction_group': employee['direction'], 'call_direction': 'Исходящий' if str(row['CALL_TYPE'])=='1' else 'Входящий',
            'duration_seconds': int(row['CALL_DURATION']), 'call_result': 'Успешный',
            'crm_activity_id': str(row.get('CRM_ACTIVITY_ID') or ''), 'record_file_id': str(row.get('RECORD_FILE_ID') or '')})
        if item.get('local_path') and Path(item['local_path']).exists():
            if validate_audio(Path(item['local_path'])) != item.get('sha256'):
                raise SafeError('EXISTING_AUDIO_HASH_MISMATCH')
            continue
        if limit and processed >= limit:
            item.setdefault('status', 'Ожидает скачивания')
            continue
        processed += 1
        try:
            kind = {'CONTACT': '3', 'COMPANY': '4', 'LEAD': '1', 'DEAL': '2'}.get(row.get('CRM_ENTITY_TYPE'))
            item['crm_url'] = crm_link(kind, row.get('CRM_ENTITY_ID')) if kind else ''
            if row.get('CRM_ACTIVITY_ID'):
                try:
                    activity = read_call(client, 'crm.activity.get', {'id': int(row['CRM_ACTIVITY_ID'])})['result']
                    item['crm_url'] = crm_link(activity.get('OWNER_TYPE_ID'), activity.get('OWNER_ID')) or item['crm_url']
                    item['crm_lookup_error'] = ''
                except SafeError as error:
                    item['crm_lookup_error'] = str(error)
            if not row.get('RECORD_FILE_ID'):
                raise SafeError('NO_RECORD_FILE')
            record = read_call(client, 'disk.file.get', {'id': int(row['RECORD_FILE_ID'])})['result']
            download_url = record.get('DOWNLOAD_URL', '')
            parsed = urlsplit(download_url)
            if parsed.scheme != 'https' or parsed.hostname != PORTAL_HOST or parsed.username or parsed.password:
                raise SafeError('RECORD_URL_NOT_EXPECTED_PORTAL')
            ext = Path(record.get('NAME', '')).suffix.lower()
            if ext not in {'.mp3', '.wav', '.ogg', '.m4a', '.flac', '.aac'}:
                raise SafeError('UNSUPPORTED_AUDIO_EXTENSION')
            filename = f"{day}_{when.strftime('%H-%M-%S')}_{employee['name']}_{call_id}{ext}"
            filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
            target = out / filename
            temp = out / (filename + '.part')
            opener = urllib.request.build_opener(AudioRedirect())
            with opener.open(urllib.request.Request(download_url), timeout=45) as response, temp.open('wb') as stream:
                if 'text/html' in response.headers.get('Content-Type', ''):
                    raise SafeError('AUDIO_RESPONSE_IS_HTML')
                count = 0
                while chunk := response.read(1024*1024):
                    count += len(chunk)
                    if count > 250_000_000:
                        raise SafeError('AUDIO_TOO_LARGE')
                    stream.write(chunk)
            if record.get('SIZE') and temp.stat().st_size != int(record['SIZE']):
                raise SafeError('AUDIO_SIZE_MISMATCH')
            sha = validate_audio(temp)
            temp.replace(target)
            item.update({'local_path': str(target), 'filename': filename, 'size': target.stat().st_size,
                         'sha256': sha, 'mime_type': {'.mp3':'audio/mpeg','.wav':'audio/wav','.ogg':'audio/ogg','.m4a':'audio/mp4','.flac':'audio/flac','.aac':'audio/aac'}[ext],
                         'status': 'Скачано локально, ожидает загрузки на Drive', 'error': ''})
        except SafeError as error:
            item.update({'status': 'Ошибка', 'error': str(error)})
        except Exception:
            item.update({'status': 'Ошибка', 'error': 'DOWNLOAD_OR_NETWORK_ERROR'})
        save_json(manifest_path, manifest)
        print(json.dumps({'call_id': call_id, 'employee': employee['name'], 'status': item['status'], 'error': item.get('error', '')}, ensure_ascii=False), flush=True)
    save_json(manifest_path, manifest)
    print(json.dumps({'manifest': str(manifest_path), 'eligible': len(rows),
        'local_files': sum(bool(c.get('local_path')) for c in manifest['calls'].values()),
        'errors': sum(c.get('status')=='Ошибка' for c in manifest['calls'].values())}, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    args = argparse.ArgumentParser()
    args.add_argument('--date', default=(datetime.now(MSK)-timedelta(days=1)).date().isoformat())
    args.add_argument('--limit', type=int, default=0)
    args.add_argument('--allow-partial-mapping', action='store_true')
    options = args.parse_args()
    try:
        run(options.date, options.limit, options.allow_partial_mapping)
    except SafeError as error:
        print(json.dumps({'safe_error': str(error)}, ensure_ascii=False))
        sys.exit(1)
    except Exception:
        print(json.dumps({'safe_error': 'LOCAL_COLLECTOR_ERROR'}))
        sys.exit(1)
