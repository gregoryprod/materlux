"""Read existing BitrixGPT transcripts; never trigger AI or change CRM."""
import sys, json, re, urllib.request
from urllib.error import HTTPError
from datetime import datetime, timezone
from download_calls import ROOT, save_json
from read_check import Client

def run(day):
    path = ROOT/day/'Журнал выгрузки.json'
    data = json.loads(path.read_text(encoding='utf-8'))
    client = Client()
    counts = {}
    for call in data['calls'].values():
        if call.get('transcript_text'):
            status = 'Готово'
        else:
            status = 'Нет доступной расшифровки'
            activity = call.get('crm_activity_id')
            if str(activity or '').isdigit() and int(activity)>0:
                try:
                    req = urllib.request.Request(client._base+'crm.activity.call.getTranscript.json', data=json.dumps({'activityId':int(activity)}).encode(), headers={'Content-Type':'application/json'}, method='POST')
                    try:
                        with client._opener.open(req, timeout=25) as response:
                            result = json.loads(response.read(4000000))
                    except HTTPError as error:
                        result = json.loads(error.read(8192))
                    if result.get('error'):
                        code = str(result['error'])
                        call['transcript_error'] = code if re.fullmatch(r'[A-Za-z0-9_.-]{1,80}',code) else 'API_ERROR'
                        status = 'Ошибка доступа или API'
                    else:
                        value = result.get('result')
                        value = value.get('transcription') if isinstance(value,dict) else None
                        if isinstance(value,str) and value.strip():
                            call['transcript_text'] = value
                            call['transcript_error'] = ''
                            status = 'Готово'
                except Exception:
                    call['transcript_error'] = 'NETWORK_OR_RESPONSE_ERROR'
                    status = 'Ошибка доступа или API'
            else:
                status = 'Нет ID CRM-дела'
                call['transcript_error'] = ''
        call['transcript_status'] = status
        call['transcript_checked_at'] = datetime.now(timezone.utc).isoformat()
        counts[status] = counts.get(status,0)+1
        save_json(path,data)
    print(json.dumps({'day':day,'counts':counts},ensure_ascii=True))

if __name__ == '__main__':
    run(sys.argv[1])
