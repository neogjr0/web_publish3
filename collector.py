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

# 부동산 타입별 API 설정
PROPERTY_TYPES = {
    'apt': {
        'label': '아파트',
        'url': 'http://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade',
        'name_field': 'aptNm',
        'area_field': 'excluUseAr',
    },
    'offi': {
        'label': '오피스텔',
        'url': 'http://apis.data.go.kr/1613000/RTMSDataSvcOffiTrade/getRTMSDataSvcOffiTrade',
        'name_field': 'offiNm',
        'area_field': 'excluUseAr',
    },
    'rh': {
        'label': '연립다세대',
        'url': 'http://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade',
        'name_field': 'mhouseNm',
        'area_field': 'excluUseAr',
    },
}

def get_target_months():
    months = []
    current_date = datetime.datetime.now()
    for i in range(3):
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
    with open('setup.py', 'w', encoding='utf-8') as f:
        f.write(setup_content)


def create_html(data_list):
    json_data = json.dumps(data_list, ensure_ascii=False)
    district_keys = list(DISTRICTS.keys())

    html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SEOUL REAL - 프리미엄 리포트</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;700;900&display=swap');
        body {{ background: #020617; color: #ffffff; font-family: 'Pretendard', sans-serif; }}
        .glass {{ background: #0f172a; border: 1px solid #1e293b; border-radius: 1rem; }}
        .filter-btn {{ background: #1e293b; border: 1px solid #334155; padding: 10px 16px; border-radius: 10px; font-weight: 800; color: #94a3b8; transition: 0.2s; font-size: 0.85rem; cursor: pointer; }}
        .filter-btn:hover {{ border-color: #475569; color: #cbd5e1; }}
        .filter-btn.active {{ background: #3b82f6; color: #fff; border-color: #3b82f6; }}
        /* 타입 필터 버튼 */
        .type-btn {{ background: #1e293b; border: 1px solid #334155; padding: 8px 16px; border-radius: 8px; font-weight: 800; color: #94a3b8; transition: 0.2s; font-size: 0.8rem; cursor: pointer; }}
        .type-btn:hover {{ border-color: #475569; }}
        .type-btn.active-apt {{ background: #3b82f6; color: #fff; border-color: #3b82f6; }}
        .type-btn.active-offi {{ background: #8b5cf6; color: #fff; border-color: #8b5cf6; }}
        .type-btn.active-rh {{ background: #10b981; color: #fff; border-color: #10b981; }}
        .type-btn.active-all {{ background: #f59e0b; color: #fff; border-color: #f59e0b; }}
        .download-btn {{ background: #059669; color: #fff; padding: 10px 18px; border-radius: 10px; font-weight: 800; font-size: 0.9rem; transition: 0.2s; cursor: pointer; }}
        .download-btn:hover {{ background: #10b981; transform: scale(1.05); }}
        .apt-card {{ transition: all 0.3s; border: 2px solid #1e293b; cursor: pointer; }}
        .apt-card:hover {{ border-color: #3b82f6; background: #111d35; transform: translateY(-5px); }}
        .apt-card.type-offi:hover {{ border-color: #8b5cf6; }}
        .apt-card.type-rh:hover {{ border-color: #10b981; }}
        .text-price {{ color: #60a5fa; font-size: 1.8rem; font-weight: 900; }}
        .text-date {{ background: #3b82f6; padding: 2px 8px; border-radius: 4px; font-size: 0.8rem; font-weight: 800; }}
        .badge-apt {{ background: #1d4ed8; color: #93c5fd; padding: 2px 8px; border-radius: 4px; font-size: 0.7rem; font-weight: 800; }}
        .badge-offi {{ background: #4c1d95; color: #c4b5fd; padding: 2px 8px; border-radius: 4px; font-size: 0.7rem; font-weight: 800; }}
        .badge-rh {{ background: #064e3b; color: #6ee7b7; padding: 2px 8px; border-radius: 4px; font-size: 0.7rem; font-weight: 800; }}
        @media print {{ .filter-btn, .type-btn, .download-btn, header button {{ display: none !important; }} }}
    </style>
</head>
<body class="p-4 md:p-10">
    <div class="max-w-[1600px] mx-auto">
        <header class="flex flex-wrap justify-between items-center mb-10 border-b border-slate-800 pb-8 gap-4">
            <div>
                <h1 class="text-4xl font-black italic mb-2">SEOUL <span class="text-blue-500">REAL</span></h1>
                <p class="text-2xl font-black text-white tracking-tight">서울 전지역 최근 3개월 실거래 조회 서비스</p>
                <p class="text-slate-500 text-sm mt-1">아파트 · 오피스텔 · 연립다세대 포함</p>
            </div>
            <div class="flex gap-3 flex-wrap">
                <button onclick="downloadExcel()" class="download-btn">📗 Excel 다운로드</button>
                <button onclick="window.print()" class="download-btn" style="background:#dc2626;">📕 PDF 다운로드</button>
                <button onclick="location.reload()" class="bg-slate-800 px-4 py-2 rounded-lg font-bold text-sm cursor-pointer">🔄 초기화</button>
            </div>
        </header>

        <!-- 부동산 타입 필터 -->
        <div class="glass p-4 mb-4">
            <p class="text-xs font-bold text-slate-500 mb-3 uppercase tracking-widest">🏢 부동산 유형</p>
            <div class="flex flex-wrap gap-2">
                <button class="type-btn active-all" onclick="filterByType('all', this)">전체</button>
                <button class="type-btn" onclick="filterByType('apt', this)">🏠 아파트</button>
                <button class="type-btn" onclick="filterByType('offi', this)">🏢 오피스텔</button>
                <button class="type-btn" onclick="filterByType('rh', this)">🏘️ 연립·다세대</button>
            </div>
        </div>

        <!-- 구 필터 -->
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
                <h2 class="text-sm font-bold text-emerald-400 mb-4 uppercase tracking-widest">🏘️ 유형별 거래 비중</h2>
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
        const districts = {district_keys};
        let charts = {{}};
        let currentTypeFilter = 'all';
        let currentGuFilter = '전체';

        const TYPE_LABELS = {{ apt: '아파트', offi: '오피스텔', rh: '연립·다세대' }};
        const TYPE_COLORS = {{ apt: '#3b82f6', offi: '#8b5cf6', rh: '#10b981' }};

        function init() {{
            const group = document.getElementById('filterGroup');
            districts.forEach(gu => {{
                const btn = document.createElement('button');
                btn.className = 'filter-btn';
                btn.innerText = gu;
                btn.onclick = (e) => filterByGu(gu, e.target);
                group.appendChild(btn);
            }});
            applyFilters();
        }}

        function filterByType(type, btn) {{
            currentTypeFilter = type;
            document.querySelectorAll('.type-btn').forEach(b => {{
                b.className = 'type-btn';
            }});
            const activeClass = type === 'all' ? 'active-all' : `active-${{type}}`;
            btn.classList.add(activeClass);
            applyFilters();
        }}

        function filterByGu(gu, btn) {{
            currentGuFilter = gu;
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            if(btn) btn.classList.add('active');
            applyFilters();
        }}

        function applyFilters() {{
            let data = rawData;
            if (currentTypeFilter !== 'all') {{
                data = data.filter(d => d.type === currentTypeFilter);
            }}
            if (currentGuFilter !== '전체') {{
                data = data.filter(d => d.gu === currentGuFilter);
            }}
            renderDashboard(data, currentGuFilter);
        }}

        function renderDashboard(data, label) {{
            window.currentFilteredData = data;
            document.getElementById('totalCount').innerText = data.length.toLocaleString();
            document.getElementById('selectedLabel').innerText = label;

            const listBody = document.getElementById('dataList');
            listBody.innerHTML = data.slice(0, 300).map(item => {{
                const p = item.price;
                const eok = Math.floor(p / 10000);
                const man = p % 10000;
                const priceStr = eok > 0
                    ? (eok + '억 ' + (man > 0 ? man.toLocaleString() + '만' : ''))
                    : p.toLocaleString() + '만';
                const searchUrl = `https://m.land.naver.com/search/result/${{item.gu}} ${{item.umd}} ${{item.aptNm}}`;
                const badgeClass = item.type === 'offi' ? 'badge-offi' : item.type === 'rh' ? 'badge-rh' : 'badge-apt';
                const typeLabel = TYPE_LABELS[item.type] || '아파트';
                const cardClass = item.type === 'offi' ? 'type-offi' : item.type === 'rh' ? 'type-rh' : '';

                return `
                <div class="glass p-6 apt-card ${{cardClass}}" onclick="window.open('${{searchUrl}}', '_blank')">
                    <div class="flex justify-between items-start mb-3">
                        <div class="flex gap-2 items-center">
                            <span class="text-xs font-bold text-blue-400">${{item.gu}} ${{item.umd}}</span>
                            <span class="${{badgeClass}}">${{typeLabel}}</span>
                        </div>
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

            // 유형별 집계
            const typeCount = {{ apt: 0, offi: 0, rh: 0 }};

            data.forEach(d => {{
                const key = (label === '전체') ? d.gu : d.umd;
                countMap[key] = (countMap[key] || 0) + 1;
                priceMap[key] = (priceMap[key] || []);
                priceMap[key].push(d.price);
                if (typeCount[d.type] !== undefined) typeCount[d.type]++;
            }});

            const labels = Object.keys(countMap);
            const avgPrices = labels.map(k => {{
                const v = priceMap[k];
                return Math.round(v.reduce((a, b) => a + b, 0) / v.length);
            }});

            renderChart('mainChart', 'bar', labels, [{{
                label: '거래건수', data: Object.values(countMap), backgroundColor: '#3b82f6', borderRadius: 5
            }}], true);

            renderChart('priceLineChart', 'line', labels, [{{
                label: '평균가', data: avgPrices, borderColor: '#60a5fa', tension: 0.3, fill: true,
                backgroundColor: 'rgba(59, 130, 246, 0.1)'
            }}]);

            renderChart('areaPieChart', 'doughnut',
                Object.keys(typeCount).map(k => TYPE_LABELS[k]),
                [{{
                    data: Object.values(typeCount),
                    backgroundColor: ['#3b82f6', '#8b5cf6', '#10b981'],
                    borderWidth: 0
                }}]
            );
        }}

        function renderChart(id, type, labels, datasets, clickable = false) {{
            if (charts[id]) charts[id].destroy();
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
                            if (currentRegion === '전체') {{
                                // 구 버튼 찾아서 클릭 효과
                                const btns = document.querySelectorAll('.filter-btn');
                                btns.forEach(b => {{ if (b.innerText === clickedLabel) filterByGu(clickedLabel, b); }});
                            }} else {{
                                const filtered = window.currentFilteredData.filter(d => d.umd === clickedLabel);
                                renderDashboard(filtered, clickedLabel);
                            }}
                        }}
                    }} : null,
                    plugins: {{
                        legend: {{
                            display: type === 'doughnut', position: 'bottom',
                            labels: {{ color: '#94a3b8', padding: 12 }}
                        }}
                    }},
                    scales: type !== 'doughnut' ? {{
                        y: {{ grid: {{ color: '#1e293b' }}, ticks: {{ color: '#64748b' }} }},
                        x: {{ grid: {{ display: false }}, ticks: {{ color: '#fff', font: {{ size: 10, weight: 'bold' }} }} }}
                    }} : {{}}
                }}
            }});
        }}

        function downloadExcel() {{
            const exportData = window.currentFilteredData.map(item => ({{
                '유형': TYPE_LABELS[item.type] || item.type,
                '구': item.gu, '동': item.umd,
                '건물명': item.aptNm,
                '금액(만원)': item.price,
                '전용면적(㎡)': item.area,
                '거래일': item.date
            }}));
            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "RealEstate");
            XLSX.writeFile(wb, `SEOUL_REAL_3MONTHS_${{new Date().getTime()}}.xlsx`);
        }}

        init();
    </script>
</body>
</html>"""

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)


def collect():
    all_data = []
    target_months = get_target_months()
    print(f"🚀 {target_months} 수집 시작...")

    for name, code in DISTRICTS.items():
        for ptype, cfg in PROPERTY_TYPES.items():
            for month in target_months:
                params = {
                    'serviceKey': unquote(API_KEY),
                    'LAWD_CD': code,
                    'DEAL_YMD': month,
                    'numOfRows': '500'
                }
                try:
                    res = requests.get(cfg['url'], params=params, timeout=10)
                    items = ET.fromstring(res.text).findall('.//item')
                    for item in items:
                        y = item.findtext('dealYear', '').strip()
                        m = (item.findtext('dealMonth') or '').strip().zfill(2)
                        d = (item.findtext('dealDay') or '').strip().zfill(2)
                        name_val = item.findtext(cfg['name_field'], '').strip()
                        area_val = item.findtext(cfg['area_field'], '0').strip()
                        price_val = item.findtext('dealAmount', '0').replace(',', '').strip()

                        if not name_val or not price_val:
                            continue

                        all_data.append({
                            'type': ptype,
                            'gu': name,
                            'umd': item.findtext('umdNm', '').strip(),
                            'aptNm': name_val,
                            'price': int(price_val),
                            'area': float(area_val) if area_val else 0.0,
                            'date': f"{y}-{m}-{d}"
                        })
                except Exception as e:
                    pass

        print(f"✔️ {name} 완료 (아파트+오피스텔+연립다세대)")

    all_data.sort(key=lambda x: x['date'], reverse=True)
    create_html(all_data)
    create_automation_files()
    print(f"✨ 완료! 총 {len(all_data)}건 수집")


if __name__ == "__main__":
    collect()
