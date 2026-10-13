"""Meaningful regression tests for the independent auditor; source files stay read-only."""
import argparse
import copy
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--manifest',type=Path,required=True)
    p.add_argument('--node',required=True)
    p.add_argument('--scratch-root',type=Path,default=Path.cwd(),help='Разрешённая рабочая папка для временных копий')
    args=p.parse_args()
    manifest=json.loads(args.manifest.read_text(encoding='utf-8'))
    original=Path(manifest['dashboard']).read_text(encoding='utf-8')
    pattern=r'(<script id="payload" type="application/json">)(.*?)(</script>)'
    data=json.loads(re.search(pattern,original,re.S)[2])
    script=Path(__file__).with_name('critical_check.py')
    def invoke(path):
        result=subprocess.run([sys.executable,str(script),'--manifest',str(path),'--node',args.node],capture_output=True,text=True,encoding='utf-8',timeout=60)
        return result.returncode,json.loads(result.stdout)
    cases=[
        ('baseline',None,None,True),
        ('quantity_changed','qty',None,False),
        ('width_boundary_120','width',None,False),
        ('wrong_order_address','address',None,False),
        ('wrong_city_tariff','tariff',None,False),
        ('excluded_region_as_free','excluded',None,False),
        ('october_source_missing',None,'october',False),
        ('pickup_instead_of_door','mode',None,False),
        ('javascript_shipping_total_wrong',None,'javascript',False),
        ('hidden_cost_sheet',None,'hidden',False),
    ]
    results=[]
    scratch_root=args.scratch_root.resolve()
    assert scratch_root.is_dir()
    tmp=tempfile.TemporaryDirectory(prefix='mnogo_sna_skill_test_',dir=str(scratch_root))
    temp_root=Path(tmp.name).resolve()
    # Confirm the exact cleanup target lies within the OS temporary root.
    assert temp_root.parent==scratch_root
    try:
        for name,mutation,extra,expected_pass in cases:
            test=copy.deepcopy(data);m=copy.deepcopy(manifest);html=original
            if mutation=='qty':test['items'][0]['qty']+=1
            if mutation=='width':next(x for x in test['items'] if x['width']==120)['width']=121
            if mutation=='address':test['items'][0]['address']=test['items'][3]['address']
            if mutation=='tariff':next(x for x in test['items'] if x['shipping_unit'] is not None)['shipping_unit']+=10
            if mutation=='excluded':
                row=next(x for x in test['items'] if x['region']=='Москва и Московская область');row['shipping_unit']=0;row['shipping_total']=0
            if mutation=='mode':test['shipping']['destination_mode']='Получить груз в отделении'
            if extra=='october':m['orders']=m['orders'][:1]
            if extra=='javascript':
                assert 'a.shippingTotal+=x.shipping_total' in html
                html=html.replace('a.shippingTotal+=x.shipping_total','a.shippingTotal+=0')
            if extra=='hidden':
                fixture=temp_root/'hidden_fixture.xlsx'
                book=openpyxl.Workbook();book.active.title='Visible fixture';hidden=book.create_sheet('Napoli');hidden.sheet_state='hidden';book.save(fixture)
                m['price']['cost_source']=str(fixture)
            rendered=re.sub(pattern,lambda match:match[1]+json.dumps(test,ensure_ascii=False).replace('<','\\u003c')+match[3],html,flags=re.S)
            output=temp_root/(name+'.html');output.write_text(rendered,encoding='utf-8');m['dashboard']=str(output)
            config=temp_root/(name+'.json');config.write_text(json.dumps(m,ensure_ascii=False),encoding='utf-8')
            code,report=invoke(config)
            passed=code==0 and report['status']=='NUMERIC_PASS'
            assert passed==expected_pass,(name,code,report)
            failures=[x['name'] for x in report['checks'] if x['status']=='FAIL']
            results.append({'case':name,'expected': 'PASS' if expected_pass else 'FAIL','observed':report['status'],'detected_by':failures})
    finally:
        assert temp_root.parent==scratch_root
        tmp.cleanup()
    print(json.dumps({'tests':len(results),'passed':True,'results':results},ensure_ascii=False,indent=2))


if __name__=='__main__':
    main()
