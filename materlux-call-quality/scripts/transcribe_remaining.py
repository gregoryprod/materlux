"""Offline, resumable Whisper medium fallback for recordings without text."""
import os, sys, json, time, hashlib
from pathlib import Path
from datetime import datetime, timezone
BASE=Path(__file__).resolve().parent
# Dependencies installed in the selected Python environment.
os.environ['HF_HUB_OFFLINE']='1'
os.environ['HF_HUB_DISABLE_TELEMETRY']='1'
from faster_whisper import WhisperModel, BatchedInferencePipeline
from download_calls import ROOT, save_json

day=sys.argv[1]
path=ROOT/day/'Журнал выгрузки.json'
data=json.loads(path.read_text(encoding='utf-8'))
out=ROOT/day/'Whisper medium'
out.mkdir(exist_ok=True)
selected=[c for c in data['calls'].values() if not c.get('transcript_text')]
print(json.dumps({'selected':len(selected),'audio_seconds':sum(c['duration_seconds'] for c in selected)}),flush=True)
model=WhisperModel(os.environ['MATERLUX_MODEL_PATH'],device='cpu',compute_type='int8',cpu_threads=4,num_workers=1,local_files_only=True)
if '--batched' in sys.argv:model=BatchedInferencePipeline(model=model)
for c in selected:
    try:
        source=Path(c['local_path'])
        assert hashlib.sha256(source.read_bytes()).hexdigest()==c['sha256']
        started=time.perf_counter()
        options=dict(language='ru',beam_size=5,vad_filter=True,word_timestamps='--segment-only' not in sys.argv,condition_on_previous_text=False)
        if '--batched' in sys.argv:options['batch_size']=4
        segments,info=model.transcribe(str(source),**options)
        items=[]
        for s in segments:
            items.append({'start':s.start,'end':s.end,'text':s.text.strip(),'avg_logprob':s.avg_logprob,'no_speech_prob':s.no_speech_prob,'words':[{'start':w.start,'end':w.end,'word':w.word,'probability':w.probability} for w in (s.words or [])]})
        text='\n'.join(f'[{int(s["start"])//60:02d}:{int(s["start"])%60:02d}–{int(s["end"])//60:02d}:{int(s["end"])%60:02d}] {s["text"]}' for s in items)
        result={'call_id':c['call_id'],'source_sha256':c['sha256'],'model':'Whisper medium','segments':items,'decoded_duration':info.duration,'processing_seconds':round(time.perf_counter()-started,2),'quality_status':'Автоматический текст, точность по аудио не подтверждена'}
        save_json(out/(c['call_id']+'.json'),result)
        (out/(c['call_id']+'.txt')).write_text(text,encoding='utf-8')
        if text.strip():
            c.update(transcript_text=text,transcript_status='Готово',transcript_source='Whisper medium (локально)',transcript_quality='Требует проверки по аудио',transcript_error='',transcript_checked_at=datetime.now(timezone.utc).isoformat())
        else:
            c['transcript_status']='Whisper: речь не распознана'
        save_json(path,data)
        print(json.dumps({'call_id':c['call_id'],'characters':len(text),'seconds':result['processing_seconds'],'completed':sum(bool(x.get('transcript_text')) for x in data['calls'].values())}),flush=True)
    except Exception:
        c['transcript_status']='Ошибка локальной расшифровки'
        save_json(path,data)
        print(json.dumps({'call_id':c['call_id'],'error':'LOCAL_TRANSCRIPTION_ERROR'}),flush=True)
