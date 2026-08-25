# -*- coding: utf-8 -*-
"""최종 빌드 : s2b.xlsx → final4.xlsx (분석 10개 시트)"""
exec(open('fbase.py').read())
for f in ['f_macro.py','f_struct.py','f_sec.py','f_chain.py','f_hs4.py','f_margin.py','f_idea.py','f_chart.py','f_sum.py','f_guide.py']:
    exec(open(f).read())
ORDER=['0_안내','분석1_핵심요약','분석2_월별거시','분석3_수출구조분해','분석4_섹터스코어보드','분석5_섹터월별',
       '분석6_산업체인','분석7_HS4상세','분석8_마진·단가','분석9_스크리너·투자아이디어','분석10_그래프',
       '수출입 총괄','수출입 실적(품목별)','수출입 실적(성질별_수출)','수출입 실적(성질별_수입)','수출입 실적(신성질별)',
       'HS4_85_87_90','HS4_84','HS4_27_29_71','HS4_39_72_89',
       '데이터_총괄','데이터_품목별','데이터_성질별수출','데이터_성질별수입','데이터_신성질별',
       '데이터_HS4수출','데이터_HS4수입','데이터_HS4중량','데이터_HS4수입중량']
missing=[s for s in wb.sheetnames if s not in ORDER]
assert not missing, missing
wb._sheets=[wb[s] for s in ORDER]
wb.active=0
wb.save('final4.xlsx')
print('saved final4 :', len(wb.sheetnames),'시트')
