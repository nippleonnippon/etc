import Parser from 'rss-parser';
import fs from 'fs';
import XLSX from 'xlsx';

// ✅ 設定パラメータ
const DEFAULT_ITEMS_PER_PAGE = 100;   // 1ページあたりの表示件数
const DEFAULT_MAX_ITEMS = 2000;       // 最大保存件数
const DEFAULT_NAV_GROUP_SIZE = 5;    // ページナビゲーションの1グループ表示数
const fetchMinuit = 3;

// ExcelファイルからRSS情報を読み込む関数
function loadRssUrlsFromExcel(filepath) {
  const workbook = XLSX.readFile(filepath);
  const sheetName = workbook.SheetNames[0]; // 最初のシート
  const sheet = workbook.Sheets[sheetName];

  // シートの内容を配列に変換 (header: 1 で行ごとの配列に)
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const urls = [];
  for (let i = 1; i < data.length; i++) {  // 1行目はヘッダーなのでスキップ
    const row = data[i];
    if (row && row[0] && row[1]) {
      urls.push({ source: row[0], url: row[1] });
    }
  }
  return urls;
}

// Excelファイル名
const excelFile = 'feeds.xlsx';

// ExcelからrssUrlsを読み込む
const rssUrls = loadRssUrlsFromExcel(excelFile);

const parser = new Parser({
  customFields: {
    item: [
      ['dc:date', 'dcDate']
    ]
  }
});

const parseDateSafe = (str) => {
  const d = new Date(str);
  return isNaN(d.getTime()) ? new Date(0) : d;
};

async function fetchRSS() {
  const promises = rssUrls.map(async feedInfo => {
    try {
      const feed = await parser.parseURL(feedInfo.url);
      return feed.items.map(item => {
        const rawDate = item.pubDate || item.dcDate || '';
        const pubDate = parseDateSafe(rawDate);
        return {
          title: item.title || '',
          link: item.link || '',
          pubDate,
          source: feedInfo.source || feed.title || ''
        };
      });
    } catch (error) {
      console.error(`❌ Error fetching ${feedInfo.url}:`, error);
      return [];
    }
  });

  let allItems = (await Promise.all(promises)).flat();

  allItems.sort((a, b) => b.pubDate - a.pubDate);
  if (allItems.length > DEFAULT_MAX_ITEMS) allItems = allItems.slice(0, DEFAULT_MAX_ITEMS);


  const itemsPerPage = DEFAULT_ITEMS_PER_PAGE;
  const totalPages = Math.ceil(allItems.length / itemsPerPage);

  // ✅ ページナビゲーションの1グループ表示数も設定パラメータから取得
  const navGroupSize = DEFAULT_NAV_GROUP_SIZE;

  // ページナビのJS（表示は1ページ目固定）
  const navScript = `
<script>
  const items = ${JSON.stringify(allItems)};
  const itemsPerPage = ${itemsPerPage};
  const totalPages = ${totalPages};
  const navGroupSize = ${navGroupSize};
  let currentPage = 1;

  function formatDateJp(dateStr) {
    const d = new Date(dateStr);
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const day = d.getDate();
    const week = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
    return \`\${y}年\${m}月\${day}日（\${week}）\`;
  }

  function renderItems() {
    const container = document.querySelector('.items-container');
    container.innerHTML = '';

    const startIdx = (currentPage -1) * itemsPerPage;
    const endIdx = Math.min(startIdx + itemsPerPage, items.length);
    const pageItems = items.slice(startIdx, endIdx);

    let lastDateKey = null;

    pageItems.forEach(item => {
      const d = new Date(item.pubDate);
      const dateKey = \`\${d.getFullYear()}-\${d.getMonth()+1}-\${d.getDate()}\`;
      // 日付見出し挿入（class date, 左詰めflex）
      if (dateKey !== lastDateKey) {
        const dateDiv = document.createElement('div');
        dateDiv.className = 'date';
        dateDiv.textContent = formatDateJp(item.pubDate);
        container.appendChild(dateDiv);
        lastDateKey = dateKey;
      }

      const itemDiv = document.createElement('div');
      itemDiv.className = 'items';

      const metaDiv = document.createElement('div');
      metaDiv.className = 'metadata';

      const sourceDiv = document.createElement('div');
      sourceDiv.className = 'source';
      // ソース名は先頭11文字に制限
      sourceDiv.textContent = item.source ? item.source.slice(0, 11) : '';

      const timeDiv = document.createElement('div');
      timeDiv.className = 'time-col';
      timeDiv.textContent = d.getHours().toString().padStart(2,'0') + ':' + d.getMinutes().toString().padStart(2,'0');

      metaDiv.appendChild(sourceDiv);
      metaDiv.appendChild(timeDiv);

      const titleDiv = document.createElement('div');
      titleDiv.className = 'title';

      const a = document.createElement('a');
      a.href = item.link;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.textContent = item.title;

      titleDiv.appendChild(a);

      itemDiv.appendChild(metaDiv);
      itemDiv.appendChild(titleDiv);

      container.appendChild(itemDiv);
    });
  }

  function renderNav(navContainer) {
    navContainer.innerHTML = '';

    const totalNavGroups = Math.ceil(totalPages / navGroupSize);
    const currentNavGroup = Math.floor((currentPage - 1) / navGroupSize);

    // 左矢印（«）は現在のグループが先頭でなければ表示
    if (currentNavGroup > 0) {
      const leftSpan = document.createElement('span');
      leftSpan.className = 'nav-arrow';
      leftSpan.textContent = '«';
      leftSpan.style.cursor = 'pointer';
      leftSpan.onclick = () => {
        currentPage = currentNavGroup * navGroupSize; // 前のグループ最後のページ
        renderItems();
        renderBothNav();
        window.scrollTo({ top: 0, behavior: 'auto' });
      };
      navContainer.appendChild(leftSpan);
      navContainer.appendChild(document.createTextNode(' '));
    }

    // 現在のグループのページ番号を表示
    const startPageNum = currentNavGroup * navGroupSize + 1;
    const endPageNum = Math.min(startPageNum + navGroupSize - 1, totalPages);

    for (let i = startPageNum; i <= endPageNum; i++) {
      const pageSpan = document.createElement('span');
      pageSpan.textContent = i;
      pageSpan.style.cursor = 'pointer';
      pageSpan.style.marginRight = '4px';
      if (i === currentPage) {
        pageSpan.style.fontWeight = 'bold';
        pageSpan.style.textDecoration = 'underline';
      }
      pageSpan.onclick = () => {
        currentPage = i;
        renderItems();
        renderBothNav();
        window.scrollTo({ top: 0, behavior: 'auto' });
      };
      navContainer.appendChild(pageSpan);
    }

    // 右矢印（»）は現在のグループが最後でなければ表示
    if (currentNavGroup < totalNavGroups - 1) {
      navContainer.appendChild(document.createTextNode(' '));
      const rightSpan = document.createElement('span');
      rightSpan.className = 'nav-arrow';
      rightSpan.textContent = '»';
      rightSpan.style.cursor = 'pointer';
      rightSpan.onclick = () => {
        currentPage = (currentNavGroup + 1) * navGroupSize + 1; // 次のグループ最初のページ
        renderItems();
        renderBothNav();
        window.scrollTo({ top: 0, behavior: 'auto' });
      };
      navContainer.appendChild(rightSpan);
    }
  }

  function renderBothNav() {
    const topNav = document.querySelector('.page-navigation.top');
    const bottomNav = document.querySelector('.page-navigation.bottom');
    if (topNav) renderNav(topNav);
    if (bottomNav) renderNav(bottomNav);
  }

  function updateLastUpdateTime() {
    const now = new Date();
    const hh = now.getHours().toString().padStart(2, '0');
    const mm = now.getMinutes().toString().padStart(2, '0');
    const lastUpdateDiv = document.getElementById('last-update');
    if (lastUpdateDiv) {
      lastUpdateDiv.textContent = 'Last Update - ' + hh + ':' + mm;
    }
  }

  // 初期描画
  document.addEventListener('DOMContentLoaded', () => {
    renderItems();
    renderBothNav();
    updateLastUpdateTime();
  });
</script>
`;

  const html = `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>(๑•̀д•́๑)あんてな</title>
  <style>
    html {
      font-size: 18px; /* rem の基準を変える */
    }
    body{
        display: flex;
        justify-content: center !important;
        background-color: #e0e6fa;
    }
    #main{
        width: 100% !important;
        max-width: 800px;
        margin: 0 !important;
        padding: 300px 8px 180px 8px !important;
        box-sizing: border-box !important;
          font-family: 
            -apple-system, /* iOS/macOS の San Francisco */
            "Helvetica Neue", /* iOS/macOSの標準ゴシック */
            "Hiragino Kaku Gothic ProN", /* macOS日本語ゴシック */
            Meiryo, /* Windowsのゴシック */
            sans-serif; /* 最後の保険 */
    }
    h1 {
      height: 50px;
      color: #747474;
      display: flex;
      justify-content: space-between; /* タイトルと更新時刻を左右に配置 */
      align-items: center;
      user-select: none;  
    }

    .emoji {
      font-size: 1.5rem; /* 大きめ */
    }

    .title-text {
      font-size: 1rem; /* 普通サイズ */
      font-weight: bold;
    }

    #last-update {
      align-self: flex-end;
      margin: 0 10px 10px;
      font-size: 0.7rem;
      font-weight: bold;
      color: #000;
    }

    .items-container{
        padding: 2px;
        border:  2px solid #8cb3da;        
    }
    .items{
        display: block;
        width: 100%;
        padding: 0;
        margin: 0;
        font-size: 0.8rem;
        font-weight: bold;        
        box-sizing: border-box;
    }    
    .metadata{
        display: flex;
        flex-direction: row-reverse;
        justify-content: space-between;
        align-items: center;
        width: 100%;
        font-size: 0.6rem;
        background-color: #65a2c9;
    }
    .time-col{
        display: flex;
        justify-content: center;
        align-items: center;
        padding: 0 0 0 4px;
        font-size: 0.7rem;
        color: #fff;
    }
    .source{
        width: auto;
        padding: 0 6px 0 0 ;
        font-size: 0.7rem;
        color: #fff;
    }
    .title{
        display: block;
        margin: 0 0 3px 0;
        padding: 4px;
        background-color: #fff;
    }            
    a{
        text-decoration: none;
        color: #576cc9ff;
    }
    a:visited{
        color: #c5c5c5ff;            
    }
    /* 日付見出し */
    .date {
      display: flex;
      justify-content: flex-start;
      align-items: center;
      width: 100%;
      font-size: 0.7rem;
      font-weight: bold;
      margin-bottom: 2px;
      background-color: #ff8000;
      color: #fff;
      padding: 2px 4px;
      box-sizing: border-box;
    }
    /* ページナビ */   
    .page-navigation {
        display: flex;
        justify-content: flex-start;
        align-content: center;
        margin: 10px;
        font-size: 0.8rem;
        user-select: none;
    }    
    .page-navigation > span {
        display: flex;
        justify-content: center;
        align-items: center;
        width: 24px !important;
        height: 24px !important;
        margin: 0 2px 10px;
        font-size: 0.9rem;
        border: 1px solid #8cb3da;
        text-decoration: none !important;
        color: #0075c0;
        background-color: #fff !important;
        line-height: normal; 
    }
  </style>
</head>
<body>
  <div id="main">
    <h1>
      <div id="PageTitle">
        <span class="emoji">(๑•̀д•́๑)</span> 
        <span class="title-text">あんてな</span>
        </div>
      <div id="last-update"></div>
    </h1>
    <div class="page-navigation top"></div>
    <div class="items-container"></div>
    <div class="page-navigation bottom"></div>
    ${navScript}
  </div>
</body>
</html>
`;

  fs.writeFileSync('index.html', html, 'utf-8');
}


fetchRSS();
/*
console.log('🚀Start fetch RSS🚀');
// fetchMinuit分毎に更新
setInterval(fetchRSS, fetchMinuit * 60 * 1000);

/*

npm run start

npm run tunnel

*/