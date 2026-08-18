/* 依原始 A／B／C 需求表版面產生 .xlsx（合併格、外框、欄寬照原表）。需先載入 vendor/xlsx.bundle.js（xlsx-js-style）。 */
(function(){
  if(typeof XLSX==='undefined'){console.warn('xlsx-export: XLSX 未載入');return;}

  const thin={style:'thin',color:{rgb:'FF9AA6A0'}};
  const BORDER={top:thin,bottom:thin,left:thin,right:thin};
  const STYLE={
    L:{font:{bold:true,sz:10},alignment:{vertical:'center',wrapText:true},fill:{patternType:'solid',fgColor:{rgb:'FFEAF1ED'}},border:BORDER},
    V:{font:{sz:10},alignment:{vertical:'center',wrapText:true},border:BORDER},
    O:{font:{sz:9},alignment:{vertical:'top',wrapText:true},border:BORDER},
    T:{font:{bold:true,sz:14},alignment:{horizontal:'center',vertical:'center'},fill:{patternType:'solid',fgColor:{rgb:'FFDCE7E1'}},border:BORDER}
  };
  const BORDERONLY={border:BORDER};

  function COL(letter){return letter.charCodeAt(0)-65;}
  function txt(v){return Array.isArray(v)?v.filter(x=>String(x).trim()).join('、'):(v==null?'':String(v));}
  function has(arr,kw){return (Array.isArray(arr)?arr:[arr]).some(x=>String(x||'').includes(kw));}
  function box(kw,on,extra){return (on?'■ ':'☐ ')+kw+(on&&extra?('：'+extra):'');}

  function segsToWs(segs,colW,rowH){
    const ws={},merges=[];let maxR=0,maxC=0;
    segs.forEach(s=>{
      const c1=COL(s.c),c2=COL(s.c2||s.c),r1=s.r-1,r2=(s.r2||s.r)-1;
      const ref=XLSX.utils.encode_cell({r:r1,c:c1});
      ws[ref]={v:s.t==null?'':String(s.t),t:'s',s:STYLE[s.s]||STYLE.V};
      if(c2>c1||r2>r1)merges.push({s:{r:r1,c:c1},e:{r:r2,c:c2}});
      if(r2>maxR)maxR=r2;if(c2>maxC)maxC=c2;
    });
    for(let R=0;R<=maxR;R++)for(let C=0;C<=maxC;C++){const ref=XLSX.utils.encode_cell({r:R,c:C});if(!ws[ref])ws[ref]={v:'',t:'s',s:BORDERONLY};}
    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:maxR,c:maxC}});
    ws['!merges']=merges;
    ws['!cols']=colW;
    if(rowH)ws['!rows']=rowH;
    return ws;
  }

  // 由 {ref:text} 對照表 + 合併清單重建工作表（B／C 表用，逐格照原檔）。col B 視為標籤。
  function mapToWs(cells,mergeList,colW,valueOverrides){
    Object.assign(cells,valueOverrides||{});
    let maxC=0,maxR=0;
    Object.keys(cells).forEach(ref=>{const p=XLSX.utils.decode_cell(ref);if(p.c>maxC)maxC=p.c;if(p.r>maxR)maxR=p.r;});
    (mergeList||[]).forEach(m=>{const e=XLSX.utils.decode_range(m).e;if(e.c>maxC)maxC=e.c;if(e.r>maxR)maxR=e.r;});
    const ws={};
    for(let R=0;R<=maxR;R++)for(let C=0;C<=maxC;C++){
      const ref=XLSX.utils.encode_cell({r:R,c:C});
      const val=cells[ref];
      let style=BORDERONLY;
      if(val!==undefined){
        if(C===1)style=STYLE.L;                       // 第 B 欄＝標籤
        else if(/[o○ｏ]/.test(String(val)))style=STYLE.O; // 含選項符號
        else style=STYLE.V;
      }
      ws[ref]={v:val===undefined?'':String(val),t:'s',s:style};
    }
    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:maxR,c:maxC}});
    ws['!merges']=(mergeList||[]).map(m=>XLSX.utils.decode_range(m));
    ws['!cols']=colW;
    return ws;
  }

  /* ---------------- A 表 ---------------- */
  function buildA(c){
    const s=c.summary||{},wf=c.workflowData||{};
    const adj=Array.isArray(wf['配方溝通調整'])?wf['配方溝通調整']:[];
    const a=[adj[0]||{},adj[1]||{},adj[2]||{}];
    const 劑型=txt(s['目標劑型']);
    const seg=[
      {c:'B',r:1,t:'申請人員',s:'L'},{c:'C',r:1,t:txt(s['申請人員']||c.applicant),s:'V'},{c:'D',r:1,t:'申請日期',s:'L'},{c:'E',r:1,t:txt(s['申請日期']),s:'V'},{c:'F',r:1,t:'單號',s:'L'},{c:'G',r:1,t:txt(s['單號']||c.id),s:'V'},
      {c:'B',r:2,t:'客戶名稱',s:'L'},{c:'C',c2:'G',r:2,t:txt(s['客戶名稱']||c.customer),s:'V'},
      {c:'B',r:3,t:'產品類別',s:'L'},{c:'C',r:3,t:box('人營養補充品',has(s['產品類別'],'人營養')),s:'O'},{c:'D',c2:'E',r:3,t:box('寵物補充品',has(s['產品類別'],'寵物')),s:'O'},{c:'F',c2:'G',r:3,t:'',s:'O'},
      {c:'B',r:4,t:'目標族群',s:'L'},{c:'C',r:4,t:box('全族群',has(s['族群'],'全族群'))+'  '+box('學童',has(s['族群'],'學童')),s:'O'},{c:'D',c2:'E',r:4,t:box('男性',has(s['族群'],'男性'))+'  '+box('女性',has(s['族群'],'女性'))+'  '+box('銀髮',has(s['族群'],'銀髮')),s:'O'},{c:'F',c2:'G',r:4,t:box('特殊族群',has(s['族群'],'特殊'),txt(s['特殊族群說明'])),s:'O'},
      {c:'B',r:5,t:'主訴調理功能',s:'L'},{c:'C',c2:'G',r:5,t:'調理部位：'+txt(s['調理部位'])+(s['其他調理部位']?('（'+txt(s['其他調理部位'])+'）'):'')+'\n期望調理功能：'+txt(s['期望功能'])+(s['其他期望方向']?('（'+txt(s['其他期望方向'])+'）'):'')+'\n補充說明：'+txt(s['客戶需求']),s:'O'},
      {c:'B',r:6,t:'目標成本/產品售價',s:'L'},{c:'C',c2:'G',r:6,t:'目標成本：'+txt(s['目標成本'])+'　　產品售價：'+txt(s['產品售價'])+'\n（配方原料成本不含代工包裝）',s:'O'},
      {c:'B',r:7,t:'目標劑型',s:'L'},{c:'C',c2:'G',r:7,t:[box('粉劑',/粉/.test(劑型),txt(s['粉劑規格'])),box('膠囊',/膠囊/.test(劑型),txt(s['膠囊規格'])),box('錠劑',/錠/.test(劑型),txt(s['錠劑規格'])),box('液態',/液/.test(劑型),txt(s['液態規格'])),box('果凍',/果凍/.test(劑型),txt(s['果凍規格'])),box('其他特殊劑型',/其他/.test(劑型),txt(s['其他劑型說明']))].join('  '),s:'O'},
      {c:'B',r:8,t:'配方',s:'L'},{c:'C',c2:'G',r:8,t:[box('全營養師開發',has(s['配方方式'],'全營養師')),box('偏好/指定原料',has(s['配方方式'],'偏好')||has(s['配方方式'],'指定')),box('客供配方',has(s['配方方式'],'客供')),box('仿樣',has(s['配方方式'],'仿樣'))].join('  ')+(s['配方說明']?('\n說明：'+txt(s['配方說明'])):''),s:'O'},
      {c:'B',r:9,t:'特殊需求',s:'L'},{c:'C',c2:'G',r:9,t:[box('全素',has(s['特殊需求'],'全素')),box('奶素',has(s['特殊需求'],'奶素')),box('蛋奶素',has(s['特殊需求'],'蛋奶素')),box('無香料',has(s['特殊需求'],'無香料')),box('無人工甜味劑',has(s['特殊需求'],'無人工甜味劑')),box('不額外添加',has(s['特殊需求'],'不額外添加'),txt(s['不額外添加項目'])),box('申請HALAL',has(s['特殊需求'],'HALAL')),box('申請A.A.',has(s['特殊需求'],'A.A.')),box('申請查驗登記',has(s['特殊需求'],'查驗登記')),box('輸出（出口他國販售）',has(s['特殊需求'],'輸出')||has(s['特殊需求'],'出口'),txt(s['出口國別'])),box('授權',has(s['特殊需求'],'授權')),box('專利原料',has(s['特殊需求'],'專利原料')),box('需使用專利證書/號',has(s['特殊需求'],'專利證書')),box('需有原廠研究',has(s['特殊需求'],'原廠研究'))].join('  ')+(s['其他特殊需求']?('\n其他：'+txt(s['其他特殊需求'])):''),s:'O'},
      {c:'B',r:10,t:'檢附資料',s:'L'},{c:'C',c2:'G',r:10,t:txt(s['檢附資料']),s:'V'},
      {c:'B',r:11,t:'收案日期',s:'L'},{c:'C',c2:'D',r:11,t:txt(wf['收案日期']),s:'V'},{c:'E',r:11,t:'收案營養師',s:'L'},{c:'F',c2:'G',r:11,t:txt(wf['收案營養師']),s:'V'},
      {c:'B',r:12,t:'營養師回復',s:'L'},{c:'C',c2:'G',r:12,t:txt(wf['營養師回覆']),s:'V'},
      {c:'B',r:13,t:'初步配方號碼',s:'L'},{c:'C',c2:'D',r:13,t:txt(wf['初步配方號碼']),s:'V'},{c:'E',r:13,t:'配方報價提交日期',s:'L'},{c:'F',c2:'G',r:13,t:txt(wf['配方報價提交日期']),s:'V'},
      {c:'B',r:14,r2:15,t:'配方溝通調整紀錄',s:'L'},{c:'C',c2:'D',r:14,t:'配方第一次調整',s:'L'},{c:'E',c2:'F',r:14,t:'配方第二次調整',s:'L'},{c:'G',r:14,t:'配方第三次調整',s:'L'},
      {c:'C',c2:'D',r:15,t:'日期：'+txt(a[0].date),s:'V'},{c:'E',c2:'F',r:15,t:'日期：'+txt(a[1].date),s:'V'},{c:'G',r:15,t:'日期：'+txt(a[2].date),s:'V'},
      {c:'B',r:16,t:'客戶要求',s:'L'},{c:'C',c2:'D',r:16,t:txt(a[0].request),s:'V'},{c:'E',c2:'F',r:16,t:txt(a[1].request),s:'V'},{c:'G',r:16,t:txt(a[2].request),s:'V'},
      {c:'B',r:17,t:'營養師回復',s:'L'},{c:'C',c2:'D',r:17,t:txt(a[0].reply),s:'V'},{c:'E',c2:'F',r:17,t:txt(a[1].reply),s:'V'},{c:'G',r:17,t:txt(a[2].reply),s:'V'},
      {c:'B',r:18,t:'配方號碼/提交日期',s:'L'},{c:'C',c2:'D',r:18,t:txt(a[0].formulaNo),s:'V'},{c:'E',c2:'F',r:18,t:txt(a[1].formulaNo),s:'V'},{c:'G',r:18,t:txt(a[2].formulaNo),s:'V'},
      {c:'B',r:19,r2:20,t:'打樣需求申請',s:'L'},{c:'C',c2:'G',r:19,t:'樣品超過100份 或 混和原料超過 1Kg 請改申請小量試製（收費）',s:'O'},
      {c:'C',c2:'D',r:20,t:'第一次申請',s:'L'},{c:'E',c2:'F',r:20,t:'第二次申請',s:'L'},{c:'G',r:20,t:'第三次申請',s:'L'},
      {c:'B',r:21,t:'申請日期',s:'L'},{c:'C',c2:'D',r:21,t:'',s:'V'},{c:'E',c2:'F',r:21,t:'',s:'V'},{c:'G',r:21,t:'',s:'V'},
      {c:'B',r:22,t:'樣品配方編碼',s:'L'},{c:'C',c2:'D',r:22,t:'',s:'V'},{c:'E',c2:'F',r:22,t:'',s:'V'},{c:'G',r:22,t:'',s:'V'},
      {c:'B',r:23,t:'樣品需求量',s:'L'},{c:'C',c2:'D',r:23,t:'',s:'V'},{c:'E',c2:'F',r:23,t:'',s:'V'},{c:'G',r:23,t:'',s:'V'},
      {c:'B',r:24,t:'客戶要求/調整紀錄',s:'L'},{c:'C',c2:'D',r:24,t:'',s:'V'},{c:'E',c2:'F',r:24,t:'',s:'V'},{c:'G',r:24,t:'',s:'V'},
      {c:'B',r:25,t:'研發人/收案日期',s:'L'},{c:'C',c2:'D',r:25,t:'',s:'V'},{c:'E',c2:'F',r:25,t:'',s:'V'},{c:'G',r:25,t:'',s:'V'},
      {c:'B',r:26,t:'研發回復',s:'L'},{c:'C',c2:'D',r:26,t:'',s:'V'},{c:'E',c2:'F',r:26,t:'',s:'V'},{c:'G',r:26,t:'',s:'V'},
      {c:'B',r:27,t:'樣品提交日',s:'L'},{c:'C',c2:'D',r:27,t:'',s:'V'},{c:'E',c2:'F',r:27,t:'',s:'V'},{c:'G',r:27,t:'',s:'V'},
      {c:'B',r:28,t:'客戶追蹤紀錄',s:'L'},{c:'C',c2:'G',r:28,t:txt(wf['客戶追蹤紀錄']),s:'V'},
      {c:'B',r:29,t:'是否成單',s:'L'},{c:'C',c2:'G',r:29,t:txt(wf['是否成單']),s:'V'},
      {c:'B',r:30,t:'最終確樣配方號碼',s:'L'},{c:'C',c2:'D',r:30,t:txt(wf['最終確樣配方號碼']),s:'V'},{c:'E',c2:'F',r:30,t:'結案日期',s:'L'},{c:'G',r:30,t:txt(wf['結案日期']),s:'V'}
    ];
    const colW=[{wch:3},{wch:20},{wch:16},{wch:16},{wch:22},{wch:9},{wch:30}];
    const rowH=[];rowH[4]={hpt:56};rowH[6]={hpt:44};rowH[8]={hpt:92};rowH[18]={hpt:28};
    return segsToWs(seg,colW,rowH);
  }

  /* ---------------- B 表（逐格照原檔） ---------------- */
  function buildB(c){
    const s=c.summary||{};
    const cells={
      'B1':'申請人員','D1':'申請日期','F1':'資料需求單號 :',
      'B2':'客戶名稱','F2':'配方需求單號 :',
      'B3':'產品名稱','F3':'配方編號',
      'B4':'包裝規格',
      'B5':'產品需求','C5':'o素食(全素/奶素/蛋奶素)o無香料無人工甜味劑 \no申請HALAL o申請A.A. o申請查驗登記 \no商標授權 o輸出(出口他國販售)  o其他特殊需求:',
      'B6':'彙辦單','C6':'營養標示換算單位: o最小單位-每顆/每包/其他:                       o每日劑量-         ',
      'C7':'營養標示呈現: o每日建議攝取量%  o每100g含量(大包裝建議)  o其他成分(機能成分:                        )',
      'C8':'彙辦單版本: o中文版                o中英文雙語版（用途：o出口 o哈拉 o其他:________）',
      'C9':'彙辦單提交日:','E9':'提供路徑:',
      'B10':'成分明細表','C10':'o成分明細表　＊非必要提供之資料，請確認客戶需求',
      'C11':'成分明細表版本: o中文版                o中英文雙語版（用途：o出口 o哈拉 o其他:________）',
      'C12':'成分明細表提交日:','E12':'提供路徑:',
      'B13':'產品單張','C13':'o產品單張　＊非既有提供之服務  依客戶合約需求 決定是否提供單張資料',
      'C14':'o急單　','E14':'產品主訴功能：＿＿＿＿',
      'C15':'單張提交路徑:','F15':'完成日期:',
      'B16':'商標授權','C16':'o需商標授權　＊非必要提供之資料，請確認客戶需求',
      'C17':'商標授權資料提供日：','F17':'提供路徑:',
      'B18':'包裝校搞','C19':'第一次校稿','E19':'第二次校稿','G19':'第三次校稿',
      'B20':'申請日期','B21':'包裝檔案路徑','B22':'營養師回復','B23':'提交日期',
      'B24':'商標授權','C24':'須提交資料:  1.客戶及產品中英文資料  2.授權文件(填寫用印)   3.終產品包裝(需經原廠審核)',
      'B25':'商標授權','B26':'授權原料','B27':'提交時間','B28':'資料路徑','B29':'商標授權申請',
      'B30':'包裝校正','C30':'o是 o否','E30':'o是 o否','G30':'o是 o否',
      'B31':'追蹤紀錄','B32':'完成時間','B33':'完成文件路徑','B34':'授權完成日期',
      'B35':'外包裝最終版路徑:','F35':'收件日期:',
      'B36':'認證/查登','B37':'認證/查登\n技術資料協助','C37':'需求申請日期:','E37':'申請目的:  (例如囊錠劑查驗登記)',
      'C38':'食品原料資料: oCOA(效期>1年)   o複方原料成分展開   o原料規格書(含學名及使用部位)    o進口報單',
      'C39':'o原料製程(含條件) o萃取溶劑殘留檢驗資料或聲明  o特殊原料檢疫證明文件  oHALAL證書  o其他 :',
      'C40':'食品添加物資料: oCOA(效期>1年) o複方食添成分展開含(%) o進口報單 o食添證 o產品登錄碼',
      'B41':'補件紀錄','C41':'第一次補件日期:','E41':'第二次補件日期:','G41':'第三次補件日期:',
      'B42':'補件說明備註','B43':'補件資料路徑',
      'B44':'其他需求須按件計酬＊非必要請勿提出需求申請',
      'B45':'o官網產品 審稿','D45':'是否於其他需求中提出\n行銷顧問需求單號：______________________',
      'B46':'o產品介紹PPT','B47':'o教育訓練PPT','B48':'o外出教育訓練','B49':'o其他',
      'B50':'o確認無需求，可直接歸檔','B51':'其他備註','G52':'歸檔日期: '
    };
    const merges=['B10:B12','B13:B15','B16:B17','B18:G18','B25:G25','B29:G29','B35:E35','B36:G36','B37:B40','B44:G44','B45:C45','B46:C46','B47:C47','B48:C48','B49:C49','B50:G50','B51:B52','B6:B9','C10:G10','C11:G11','C12:D12','C14:D14','C15:E15','C16:G16','C17:D17','C19:D19','C20:D20','C21:D21','C22:D22','C23:D23','C24:G24','C26:D26','C27:D27','C28:D28','C2:E2','C30:D30','C31:D31','C32:D32','C33:D33','C34:D34','C37:D37','C38:G38','C39:G39','C3:E3','C40:G40','C41:D41','C42:D42','C43:D43','C4:G4','C51:G51','C52:F52','C5:G5','C6:G6','C7:G7','C8:G8','C9:D9','D45:G49','E12:G12','E14:G14','E19:F19','E20:F20','E21:F21','E22:F22','E23:F23','E26:F26','E27:F27','E28:F28','E30:F30','E31:F31','E32:F32','E33:F33','E34:F34','E37:G37','E41:F41','E42:F42','E43:F43','E9:G9','F15:G15','F1:G1','F2:G2','F35:G35','F3:G3'];
    const colW=[{wch:3},{wch:15},{wch:26},{wch:11},{wch:26},{wch:9},{wch:30},{wch:14}];
    const overrides={'C1':txt(s['申請人員']||c.applicant),'E1':txt(s['申請日期']),'C2':txt(s['客戶名稱']||c.customer),'C3':txt(s['產品名稱']),'C4':txt(s['包裝規格'])};
    return mapToWs(cells,merges,colW,overrides);
  }

  /* ---------------- C 表（逐格照原檔） ---------------- */
  function buildC(c){
    const s=c.summary||{};
    const cells={
      'B1':'申請人員','D1':'申請日期','F1':'配方需求單號 :',
      'B2':'客戶名稱','D2':'資料路徑','F2':'行銷顧問需求單號 :C＋',
      'B3':'產品名稱','D3':'產品主訴功能',
      'B4':'配方號碼',
      'B5':'其他需求',
      'B6':'官網產品 審稿','C6':'提出時間','E6':'客戶圖片NAS路徑',
      'C7':'需求時間','E7':'客戶官網網址',
      'C8':'營養師收案時間','E8':'收案營養師','G8':'群組回覆時間',
      'C9':'第一次校稿','E9':'第二次校稿','G9':'第三次校稿',
      'B10':'申請日期','B11':'營養師回復',
      'B12':'產品介紹PPT','C12':'提出時間','E12':'產品主訴功能',
      'C13':'需求時間','E13':'收案營養師',
      'C14':'營養師收案時間','E14':'資料路徑','G14':'提供時間',
      'B15':'教育訓練PPT','C15':'o僅提供PPT','D15':'對象:','F15':'包含產品(BKE出品):',
      'D16':'客戶需求限制:','D17':'需求時間','F17':'收案營養師',
      'C18':'o須外出授課','D18':'需求時間','F18':'收案營養師',
      'D19':'總共場次：＿＿＿＿場','G19':'營養室回覆時間','H19':'最終敲定時間',
      'D20':'客戶期望時間','E20':'地點','F20':'課程時間：＿＿＿小時','G20':'日期:','H20':'日期:',
      'D21':'客戶期望時間','E21':'地點','F21':'課程時間：＿＿＿小時','G21':'日期:','H21':'日期:',
      'C22':'＊請業管與客戶溝通準備好筆電及確認當天資料完成可播放＊',
      'B23':'其他需求','C23':'SEO文','D23':'與產品無關暫不提供客戶此功能',
      'C24':'社群文案','D24':'與產品無關暫不提供客戶此功能',
      'B25':'其他備註','H25':'歸檔日期: '
    };
    const merges=['B12:B14','B15:B22','B25:B26','B5:H5','B6:B8','C10:D10','C11:D11','C15:C17','C18:C21','C22:H22','C25:G26','C4:E4','C9:D9','D15:E15','D16:E16','D17:E17','D18:E18','D19:F19','E10:F10','E11:F11','E9:F9','F15:H16','F17:H17','F18:H18','F1:H1','F2:H2','F3:H4','F6:H6','F7:H7','G10:H10','G11:H11','G9:H9','H25:H26'];
    const colW=[{wch:3},{wch:14},{wch:22},{wch:12},{wch:16},{wch:10},{wch:12},{wch:12}];
    const overrides={'C1':txt(s['申請人員']||c.applicant),'E1':txt(s['申請日期']),'C2':txt(s['客戶名稱']||c.customer),'C3':txt(s['產品名稱']),'C4':txt(s['配方號碼'])};
    return mapToWs(cells,merges,colW,overrides);
  }

  const BUILDERS={A:buildA,B:buildB,C:buildC};
  const SHEETNAME={A:'Ａ客戶配方需求表',B:'Ｂ產品資料需求',C:'Ｃ行銷顧問需求'};

  function exportCase(c){
    if(!c){return false;}
    const type=(c.type||'A').toUpperCase();
    const build=BUILDERS[type]||buildA;
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,build(c),SHEETNAME[type]||('BKE '+type));
    XLSX.writeFile(wb,`BKE_${c.id||type}_${type}表.xlsx`);
    return true;
  }

  window.BKEFormXlsx={exportCase,buildA,buildB,buildC};
})();
