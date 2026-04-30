import os, json, requests, time, datetime
import xml.etree.ElementTree as ET
from urllib.parse import unquote

# [1. 설정]
API_KEY = '7b6efd99b84e03fca06677a5f9632db682bac3e47d90f5ec37f3b4947e84307e'

DISTRICTS = {
    '종로구':'11110','중구':'11140','용산구':'11170','성동구':'11200','광진구':'11215',
    '동대문구':'11230','중랑구':'11260','성북구':'11290','강북구':'11305','도봉구':'11320',
    '노원구':'11350','은평구':'11380','서대문구':'11410','마포구':'11440','양천구':'11470',
    '강서구':'11500','구로구':'11530','금천구':'11545','영등포구':'11560','동작구':'11590',
    '관악구':'11620','서초구':'11650','강남구':'11680','송파구':'11710','강동구':'11740'
}

# [조회기간 수정: 현재 4월 기준 2, 3, 4월이 나오도록 설정]
def get_target_months():
    months = []
    # 2026년 4월 기준 -> 4월, 3월, 2월 수집
    current_date = datetime.datetime.now()
    for i in range(3):
        # 0개월 전(4월), 1개월 전(3월), 2개월 전(2월)
        target = current_date - datetime.timedelta(days=i*30)
        months.append(target.strftime("%Y%m"))
    return months

def create_automation_files():
    setup_content = """import subprocess, sys
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
if __name__ == "__main__":
    install('requests')
"""
    with open('setup.py', 'w', encoding='utf-8') as f: f.write(setup_content)

   

def create_html(data_list):
    json_data = json.dumps(data_list, ensure_ascii=False)
    html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>SEOUL REAL - 프리미엄 리포트</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;700;900&display=swap');
        body {{ background: #020617; color: #ffffff; font-family: 'Pretendard', sans-serif; }}
        .glass {{ background: #0f172a; border: 1px solid #1e293b; border-radius: 1rem; }}
        .filter-btn {{ background: #1e293b; border: 1px solid #334155; padding: 10px 16px; border-radius: 10px; font-weight: 800; color: #94a3b8; transition: 0.2s; font-size: 0.85rem; }}
        .filter-btn.active {{ background: #3b82f6; color: #fff; border-color: #3b82f6; }}
        .download-btn {{ background: #059669; color: #fff; padding: 10px 18px; border-radius: 10px; font-weight: 800; font-size: 0.9rem; transition: 0.2s; }}
        .download-btn:hover {{ background: #10b981; transform: scale(1.05); }}
        .apt-card {{ transition: all 0.3s; border: 2px solid #1e293b; cursor: pointer; }}
        .apt-card:hover {{ border-color: #3b82f6; background: #111d35; transform: translateY(-5px); }}
        .text-price {{ color: #60a5fa; font-size: 1.8rem; font-weight: 900; }}
        .text-date {{ background: #3b82f6; padding: 2px 8px; border-radius: 4px; font-size: 0.8rem; font-weight: 800; }}
        @media print {{ .filter-btn, .download-btn, header button {{ display: none !important; }} }}
    </style>
</head>
<body class="p-4 md:p-10">
    <div class="max-w-[1600px] mx-auto">
        <header class="flex flex-wrap justify-between items-center mb-10 border-b border-slate-800 pb-8 gap-4">
            <div>
                <h1 class="text-4xl font-black italic mb-2">SEOUL <span class="text-blue-500">REAL</span></h1>
                <p class="text-2xl font-black text-white tracking-tight">서울 전지역 최근 3개월 실거래 조회 서비스</p>
            </div>
            <div class="flex gap-3">
                <button onclick="downloadExcel()" class="download-btn">📗 Excel 다운로드</button>
                <button onclick="window.print()" class="download-btn" style="background:#dc2626;">📕 PDF 다운로드</button>
                <button onclick="location.reload()" class="bg-slate-800 px-4 py-2 rounded-lg font-bold text-sm">🔄 초기화</button>
            </div>
        </header>

        <div class="glass p-6 mb-10">
            <div id="filterGroup" class="flex flex-wrap gap-2">
                <button class="filter-btn active" onclick="filterByGu('전체', this)">전체보기</button>
            </div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-10">
            <div class="glass p-6">
                <h2 class="text-sm font-bold text-blue-400 mb-4 uppercase tracking-widest">📈 평균 거래가 추이 (단위: 만원)</h2>
                <div class="h-[250px]"><canvas id="priceLineChart"></canvas></div>
            </div>
            <div class="glass p-6">
                <h2 class="text-sm font-bold text-emerald-400 mb-4 uppercase tracking-widest">🏘️ 평형별 거래 비중</h2>
                <div class="h-[250px] flex justify-center"><canvas id="areaPieChart"></canvas></div>
            </div>
            <div class="glass p-6 flex flex-col justify-center items-center">
                <span id="selectedLabel" class="text-slate-500 font-bold mb-2 uppercase">Current Region</span>
                <span id="totalCount" class="text-8xl font-black text-blue-500">0</span>
                <span class="text-slate-400 font-bold mt-2">Total Transactions (3 Months)</span>
            </div>
        </div>

        <div class="glass p-8 mb-10">
            <h2 class="text-sm font-bold text-blue-400 mb-4 uppercase tracking-widest">📊 지역별 거래량 분포 (클릭 시 필터링)</h2>
            <div class="h-[300px]"><canvas id="mainChart"></canvas></div>
        </div>

        <div id="dataList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6"></div>
    </div>

    <script>
        const rawData = {json_data};
        const districts = {list(DISTRICTS.keys())};
        let charts = {{}};

        function init() {{
            const group = document.getElementById('filterGroup');
            districts.forEach(gu => {{
                const btn = document.createElement('button');
                btn.className = 'filter-btn';
                btn.innerText = gu;
                btn.onclick = (e) => filterByGu(gu, e.target);
                group.appendChild(btn);
            }});
            renderDashboard(rawData, '전체');
        }}

        function filterByGu(gu, btn) {{
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            if(btn) btn.classList.add('active');
            const filtered = gu === '전체' ? rawData : rawData.filter(d => d.gu === gu);
            renderDashboard(filtered, gu);
        }}

        function renderDashboard(data, label) {{
            window.currentFilteredData = data;
            document.getElementById('totalCount').innerText = data.length.toLocaleString();
            document.getElementById('selectedLabel').innerText = label;

            const listBody = document.getElementById('dataList');
            listBody.innerHTML = data.slice(0, 300).map(item => {{
                const p = item.price;
                const priceStr = Math.floor(p/10000) + '억 ' + (p%10000 > 0 ? (p%10000).toLocaleString() : '');
                const searchUrl = `https://m.land.naver.com/search/result/${{item.gu}} ${{item.umd}} ${{item.aptNm}}`;
                
                return `
                <div class="glass p-6 apt-card" onclick="window.open('${{searchUrl}}', '_blank')">
                    <div class="flex justify-between items-start mb-4">
                        <span class="text-xs font-bold text-blue-400">${{item.gu}} ${{item.umd}}</span>
                        <span class="text-date">${{item.date}}</span>
                    </div>
                    <h3 class="text-xl font-black mb-2 truncate">${{item.aptNm}}</h3>
                    <p class="text-slate-400 text-sm mb-6">${{item.area}}㎡ (${{Math.round(item.area/3.3)}}평)</p>
                    <div class="text-right border-t border-slate-800 pt-4"><span class="text-price">${{priceStr}}</span></div>
                </div>`;
            }}).join('');

            updateCharts(data, label);
        }}

        function updateCharts(data, label) {{
            const countMap = {{}};
            const priceMap = {{}};
            const areaMap = {{'소형(<60㎡)':0, '중형(60-85㎡)':0, '대형(>85㎡)':0}};

            data.forEach(d => {{
                const key = (label === '전체') ? d.gu : d.umd;
                countMap[key] = (countMap[key] || 0) + 1;
                priceMap[key] = (priceMap[key] || []);
                priceMap[key].push(d.price);
                
                if(d.area < 60) areaMap['소형(<60㎡)']++;
                else if(d.area <= 85) areaMap['중형(60-85㎡)']++;
                else areaMap['대형(>85㎡)']++;
            }});

            const labels = Object.keys(countMap);
            const avgPrices = labels.map(k => {{
                const v = priceMap[k];
                return Math.round(v.reduce((a,b) => a+b, 0) / v.length);
            }});

            renderChart('mainChart', 'bar', labels, [{{
                label: '거래건수', data: Object.values(countMap), backgroundColor: '#3b82f6', borderRadius: 5
            }}], true);

            renderChart('priceLineChart', 'line', labels, [{{
                label: '평균가', data: avgPrices, borderColor: '#60a5fa', tension: 0.3, fill: true, backgroundColor: 'rgba(59, 130, 246, 0.1)'
            }}]);

            renderChart('areaPieChart', 'doughnut', Object.keys(areaMap), [{{
                data: Object.values(areaMap), backgroundColor: ['#60a5fa', '#34d399', '#fbbf24'], borderWidth: 0
            }}]);
        }}

        function renderChart(id, type, labels, datasets, clickable=false) {{
            if(charts[id]) charts[id].destroy();
            const ctx = document.getElementById(id).getContext('2d');
            charts[id] = new Chart(ctx, {{
                type: type,
                data: {{ labels, datasets }},
                options: {{
                    responsive: true, maintainAspectRatio: false,
                    onClick: clickable ? (e, elements) => {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            const clickedLabel = charts[id].data.labels[idx];
                            const currentRegion = document.getElementById('selectedLabel').innerText;
                            const filtered = window.currentFilteredData.filter(d => (currentRegion === '전체' ? d.gu : d.umd) === clickedLabel);
                            renderDashboard(filtered, clickedLabel);
                        }}
                    }} : null,
                    plugins: {{ legend: {{ display: type === 'doughnut', position: 'bottom', labels: {{ color: '#94a3b8' }} }} }},
                    scales: type !== 'doughnut' ? {{
                        y: {{ grid: {{ color: '#1e293b' }}, ticks: {{ color: '#64748b' }} }},
                        x: {{ grid: {{ display: false }}, ticks: {{ color: '#fff', font: {{ size: 10, weight: 'bold' }} }} }}
                    }} : {{}}
                }}
            }});
        }}

        function downloadExcel() {{
            const exportData = window.currentFilteredData.map(item => ({{
                '구': item.gu, '동': item.umd, '아파트명': item.aptNm, '금액(만원)': item.price, '전용면적(㎡)': item.area, '거래일': item.date
            }}));
            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "RealEstate");
            XLSX.writeFile(wb, `SEOUL_REAL_3MONTHS_${{new Date().getTime()}}.xlsx`);
        }}

        init();
    </script>
</body>
</html>
    """
    with open('index.html', 'w', encoding='utf-8') as f: f.write(html_content.strip())

def collect():
    all_data = []
    target_months = get_target_months()
    print(f"🚀 {{target_months}} 수집 시작... (2, 3, 4월 데이터)".format(target_months=target_months))
    
    for name, code in DISTRICTS.items():
        for month in target_months:
            url = "http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
            params = {'serviceKey': unquote(API_KEY), 'LAWD_CD': code, 'DEAL_YMD': month, 'numOfRows': '200'}
            try:
                res = requests.get(url, params=params, timeout=10)
                items = ET.fromstring(res.text).findall('.//item')
                for item in items:
                    y = item.findtext('dealYear')
                    m = item.findtext('dealMonth').zfill(2)
                    d = item.findtext('dealDay').zfill(2)
                    all_data.append({
                        'gu': name, 'umd': item.findtext('umdNm', '').strip(),
                        'aptNm': item.findtext('aptNm').strip(),
                        'price': int(item.findtext('dealAmount').replace(',', '')),
                        'area': float(item.findtext('excluUseAr')),
                        'date': f"{{y}}-{{m}}-{{d}}".format(y=y, m=m, d=d)
                    })
            except: pass
        print(f"✔️ {{name}} 완료".format(name=name))
    
    all_data.sort(key=lambda x: x['date'], reverse=True)
    create_html(all_data)
    create_automation_files()
    print("✨ 2월부터 현재까지 3개월치 데이터 반영 완료!")

if __name__ == "__main__": collect()
