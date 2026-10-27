<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>鍵当番・金庫チェック自動割当システム</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS (xlsx) CDN -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- FontAwesome CDN -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        /* 横スクロールなしで収まるよう、表を詰めて表示する */
        .compact-table th, .compact-table td { padding-left: 0.2rem !important; padding-right: 0.2rem !important; }
        .compact-table th { min-width: 0 !important; }
        .compact-table span.inline-block { min-width: 0 !important; padding: 0.15rem 0.35rem !important; font-size: 0.75rem; }
        .compact-table select { padding: 0.15rem 0.1rem !important; font-size: 0.75rem; max-width: 100%; }
        .compact-table td:last-child { white-space: normal; }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 font-sans min-h-screen flex flex-col text-sm">
 
    <!-- ヘッダー -->
    <header class="bg-indigo-600 text-white shadow-md">
        <div class="max-w-[1920px] mx-auto px-3 py-3 flex flex-col md:flex-row justify-between items-center">
            <h1 class="text-base font-bold flex items-center gap-2">
                <i class="fa-solid fa-key"></i> 鍵当番・金庫チェック自動割当システム
            </h1>
            <p class="text-xs text-indigo-200 mt-1 md:mt-0">引渡者・引受者完全準拠版</p>
        </div>
    </header>
 
    <!-- タブナビゲーション -->
    <nav class="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-10">
        <div class="max-w-[1920px] mx-auto px-3 flex overflow-x-auto space-x-1 sm:space-x-4">
            <button type="button" onclick="switchTab('tab0')" id="btn-tab0" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-indigo-600 text-indigo-600 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-wand-magic-sparkles"></i> 0. シフト編集
            </button>
            <button type="button" onclick="switchTab('tab1')" id="btn-tab1" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-file-excel"></i> 1. シフト取込
            </button>
            <button type="button" onclick="switchTab('tab2')" id="btn-tab2" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-users-gear"></i> 2. 優先順位・ルール
            </button>
            <button type="button" onclick="switchTab('tab3')" id="btn-tab3" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-calendar-days"></i> 3. 割り当て結果
            </button>
            <button type="button" onclick="openCompareTab()" id="btn-tabCmp" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-table-columns"></i> 3-比較. 版の比較
            </button>
            <button type="button" onclick="switchTab('tab4')" id="btn-tab4" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-pen-to-square"></i> 4. 手動修正
            </button>
            <button type="button" onclick="switchTab('tab5')" id="btn-tab5" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-download"></i> 5. Excel出力
            </button>
        </div>
    </nav>
 
    <!-- メインコンテンツ -->
    <main class="max-w-[1920px] mx-auto px-3 py-6 flex-grow w-full space-y-6">
 
        <!-- 画面0: シフト編集（元のシフト表 → 編集用シフト一覧表） -->
        <section id="tab0" class="tab-content space-y-6">
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <div class="flex flex-wrap justify-between items-start gap-2 mb-2">
                    <h2 class="text-base font-bold text-slate-700 flex items-center gap-2">
                        <i class="fa-solid fa-wand-magic-sparkles text-indigo-600"></i> STEP0. 元のシフト表Excelを編集用に変換
                    </h2>
                    <button type="button" onclick="resetAll()" class="text-xs font-medium px-3 py-1.5 rounded-md border border-red-300 text-red-600 bg-white hover:bg-red-50 cursor-pointer flex items-center gap-1.5">
                        <i class="fa-solid fa-rotate-left"></i> リセット（Excelを取り除く）
                    </button>
                </div>
                <p class="text-xs sm:text-sm text-slate-500 mb-4">
                    「イオン成田_◯月シフト.xlsx」のような<strong>元のシフト表</strong>をそのまま選んでください。「【印刷】シフト」シートから、月・日数・スタッフ・シフトを自動で判定し、「◯月シフト一覧表」の形（編集用Excel）に変換して、そのまま<strong>「1. シフト取込」</strong>に入れます。11月以外の月でも使えます。
                </p>
                <input type="file" id="rawShiftFile" accept=".xlsx,.xlsm,.xls" class="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 cursor-pointer border border-slate-300 rounded-lg p-2">
                <div id="s0Result" class="hidden mt-5 space-y-3">
                    <div class="text-xs sm:text-sm font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2" id="s0Info"></div>
                    <div class="text-xs text-slate-600" id="s0Msg"></div>
                    <div class="flex flex-wrap gap-3">
                        <button type="button" onclick="s0Download()" class="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-4 py-2 rounded-lg shadow text-xs inline-flex items-center gap-2 cursor-pointer"><i class="fa-solid fa-download"></i> 編集用Excelをダウンロード (.xlsx)</button>
                        <button type="button" onclick="switchTab('tab1')" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold px-4 py-2 rounded-lg shadow text-xs inline-flex items-center gap-2 cursor-pointer">1. シフト取込へ進む <i class="fa-solid fa-arrow-right"></i></button>
                    </div>
                    <div class="overflow-auto max-h-96 border border-slate-200 rounded-lg">
                        <table class="compact-table text-xs text-center text-slate-600"><tbody id="s0Preview"></tbody></table>
                    </div>
                </div>
            </div>
        </section>

        <!-- 画面1: シフト取込 (STEP1) -->
        <section id="tab1" class="tab-content hidden space-y-6">
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <h2 class="text-base font-bold text-slate-700 mb-2 flex items-center gap-2">
                    <i class="fa-solid fa-cloud-arrow-up text-indigo-600"></i> STEP1. 編集用シフト表ファイルのアップロード
                </h2>
                <p class="text-xs sm:text-sm text-slate-500 mb-4">
                    Excelファイル（鍵管理表やシフト表）をアップロードしてください。タイトル行の年月や全日程シフトを自動で柔軟に解析します。元のシフト表は「0. シフト編集」で変換するとここに自動で入ります（変換済みの編集用Excelをここから直接選ぶこともできます）。
                </p>
                <div class="flex flex-wrap items-center gap-3 mb-4 p-3 bg-indigo-50 border border-indigo-200 rounded-lg">
                    <span class="font-bold text-indigo-700 text-xs sm:text-sm"><i class="fa-solid fa-calendar-check"></i> ① 対象の年月を選択</span>
                    <select id="selYear" class="border border-slate-300 rounded-md px-2 py-1.5 bg-white text-sm"></select>
                    <span class="text-sm">年</span>
                    <select id="selMonth" class="border border-slate-300 rounded-md px-2 py-1.5 bg-white text-sm"></select>
                    <span class="text-sm">月</span>
                    <span class="text-xs text-slate-500">② 選んだ年月のカレンダー（曜日）で処理します。続けてファイルを選択してください。</span>
                </div>
                <div class="flex flex-col sm:flex-row items-center gap-4">
                    <input type="file" id="shiftFile" accept=".xlsx, .xls, .csv" class="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 cursor-pointer border border-slate-300 rounded-lg p-2">
                </div>
            </div>
 
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="font-bold text-slate-700 flex items-center gap-2 text-sm sm:text-base">
                        <i class="fa-solid fa-table text-indigo-600"></i> STEP1. 対象メンバー・全日程シフト一覧表
                    </h3>
                    <div class="flex items-center gap-2"><span id="loadStatus" class="text-xs px-2.5 py-1 bg-amber-50 text-amber-700 rounded-full font-medium border border-amber-200">未読込（Excelを読み込んでください）</span><button type="button" onclick="resetAll()" class="text-xs font-medium px-3 py-1 rounded-md border border-red-300 text-red-600 bg-white hover:bg-red-50 cursor-pointer flex items-center gap-1.5"><i class="fa-solid fa-rotate-left"></i> リセット</button></div>
                </div>
                <div class="overflow-x-auto max-h-96 border border-slate-200 rounded-lg">
                    <table id="previewTable" class="compact-table w-full text-xs sm:text-sm text-center text-slate-600 divide-y divide-slate-200">
                        <thead class="bg-slate-100 text-slate-700 sticky top-0">
                            <tr id="previewHeadRow">
                                <th class="px-3 py-2.5 font-semibold whitespace-nowrap text-center">日付</th>
                                <th class="px-3 py-2.5 font-semibold whitespace-nowrap text-center">曜日</th>
                            </tr>
                        </thead>
                        <tbody id="previewBody" class="divide-y divide-slate-200 bg-white">
                            <tr>
                                <td colspan="7" class="px-4 py-8 text-center text-slate-400">ファイル未選択の場合、ボタンを押すと自動的にサンプルデータで作成されます。</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div class="mt-4 flex justify-end">
                    <button type="button" onclick="proceedToGenerate()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-5 py-2.5 rounded-lg shadow transition flex items-center gap-2 cursor-pointer text-sm">
                        <span>鍵当番・金庫チェックを自動作成する</span>
                        <i class="fa-solid fa-arrow-right"></i>
                    </button>
                </div>
            </div>
        </section>
 
        <!-- 画面2: 優先順位・ルール -->
        <section id="tab2" class="tab-content hidden space-y-6">
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                    <h2 class="text-base font-bold text-slate-700 mb-2 flex items-center gap-2">
                        <i class="fa-solid fa-list-ol text-indigo-600"></i> 対象者と優先順位
                    </h2>
                    <p class="text-xs sm:text-sm text-slate-500 mb-4">
                        ▲▼で優先順位を入れ替え、役職は右の入力欄で自由に書き換えられます（リーダー・店長・メンター等。候補から選ぶことも、直接入力することも可能）。
                    </p>
                    <datalist id="roleOptions">
                        <option value="リーダー"></option>
                        <option value="店長"></option>
                        <option value="副店長"></option>
                        <option value="メンター"></option>
                        <option value="スタッフ"></option>
                    </datalist>
                    <div class="space-y-2" id="priorityList"></div>

                    <div class="mt-4 p-3 border border-dashed border-indigo-300 rounded-lg bg-indigo-50/50">
                        <div class="text-xs font-bold text-indigo-700 mb-2"><i class="fa-solid fa-user-plus"></i> メンバーを追加</div>
                        <div class="flex flex-wrap items-center gap-2">
                            <input id="newStaffName" type="text" placeholder="氏名（例：田中）" class="w-28 border border-slate-300 rounded-md px-2 py-1.5 text-xs bg-white focus:outline-none focus:ring-2 focus:ring-indigo-300">
                            <input id="newStaffRole" type="text" list="roleOptions" placeholder="役職" class="w-24 border border-slate-300 rounded-md px-2 py-1.5 text-xs bg-white focus:outline-none focus:ring-2 focus:ring-indigo-300">
                            <input id="newStaffAlias" type="text" placeholder="Excel上の表記（任意）" class="w-40 border border-slate-300 rounded-md px-2 py-1.5 text-xs bg-white focus:outline-none focus:ring-2 focus:ring-indigo-300">
                            <button type="button" onclick="addStaff()" class="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-medium px-3 py-1.5 rounded-md cursor-pointer"><i class="fa-solid fa-plus"></i> 追加</button>
                        </div>
                    </div>

                    <div id="fixedHolderBox" class="mt-4 p-3 border border-slate-200 rounded-lg bg-slate-50"></div>

                    <p class="mt-3 text-[11px] text-slate-400 leading-relaxed">
                        ※氏名の変更・メンバーの追加／削除・鍵No.3担当の変更を行うと、割り当て結果は再計算されます（手動固定は解除されます）。<br>
                        ※「Excel上の表記」は、シフト表に書かれている名前（フルネーム等）です。カンマ区切りで複数登録できます。氏名を変更した場合は空欄に戻るので、必要なら入力し直してください。<br>
                        ※鍵は3つ必要なため、メンバーは最低3名・最大12名です。
                    </p>
                </div>
 
                <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 space-y-4">
                    <h2 class="text-base font-bold text-slate-700 mb-2 flex items-center gap-2">
                        <i class="fa-solid fa-shield-halved text-indigo-600"></i> 鍵管理・受渡しルール
                    </h2>
                    <div class="p-4 bg-slate-50 rounded-lg border border-slate-200 text-xs text-slate-600 space-y-2">
                        <ul class="list-disc list-inside space-y-1">
                            <li><strong>3つの鍵は必ず別々の人が保持</strong>：鍵No.1・No.2・No.3は同じ人が重複して持たないようにします（店舗置きは人ではないので除外）。</li>
                            <li><strong>鍵開けは超早・早・フルの人</strong>：前日から鍵を持っている人、または前日に誰かが店舗に置いた鍵を使える人だけが担当できます。月全体を先読みして、翌日以降の鍵開け担当に鍵が渡るよう割り当てます。休・応援・グランド等で出勤しない人は鍵を持ち帰ったまま次の出勤日に使うため、店舗に置くのは「どうしても受け渡せない時」だけです。</li>
                            <li><strong>金庫チェックは遅番の人だけ</strong>：遅番以外のシフトの人（早番・超早、会議・応援・休みなど）には金庫チェックを割り当てません。フルの人はできます。遅番がいない日は担当なしとしてNGで知らせます。画面2の「金庫チェックの担当条件」で変更できます。</li>
                            <li><strong>鍵閉め・金庫チェックは遅番・フルの人</strong>：遅番の人の中から選びます。「フル」は早番・遅番どちらにも対応可能として扱います。鍵閉め担当は退勤時に鍵を持っている必要があります。</li>
                            <li><strong><span class="fixedHolderName">五家</span>は1つの鍵を持ち回し ＆ 残り2つを分散</strong>：<span class="fixedHolderName">五家</span>が鍵No.3をずっと持ち回し、残りの鍵No.1・No.2を他のメンバーで回します。店舗置きは、持ち帰りや受け渡しで回らない時だけです。</li>
                            <li><strong>引渡者・引受者同一時の省略</strong>：引渡者と引受者が同じ場合は、引受者セルを空欄としてスッキリ表現します。</li>
                            <li><strong>鍵開けは前日の夜までに鍵を持っていること</strong>：翌日が早番（超早・早・フル）の場合、その担当者は前日の夜までに鍵を受け取っている必要があります。</li>
                            <li><strong>遅番→早番の受け渡し不可</strong>：遅番は21:30以降に閉めて帰り、早番は19:00に退勤するため、遅番から早番へ鍵を渡すことはできません（早番→遅番は可）。</li>
                            <li><strong>担当者未設定なし</strong>：鍵開け・鍵閉め・金庫チェックは必ず誰かを割当て、条件を緩めた場合は結果画面の上部にエラーとして記載します。</li>
                            <li><strong>研修の日は店舗にいない</strong>：「研修」の日は店舗にいないため、鍵の受け渡し（引渡し・引受け）・鍵開け・鍵閉め・金庫チェックはできません。その人が鍵を持っている場合は、研修の前の日までに他の人へ渡しておくか、持ち帰ったまま研修明けの出勤日に使います。</li>
                            <li><strong>会議の日は店舗に出勤しない</strong>：「会議」（店長会議・ミーティング等）の日は店舗に出勤していないため、鍵の受け渡し・鍵開け・鍵閉め・金庫チェックはできません。その人の鍵は持ち帰ったままになるので、会議の前の日までに他の人へ渡しておきます。表記はルール設定の「出勤しない表記」で追加できます。</li>
                            <li><strong>非番日考慮</strong>：応援・休・グランド・店長会議・方針等の出勤しない日では鍵の受け渡しができないため、出勤しているメンバーまたは店舗置きで適切に対応します。</li>
                            <li><strong>上位者の鍵は先回りで他の人へ</strong>：優先順位の上位2名（既定：<span class="deferTopNames">五家・爲貝</span>）には鍵を持たせるのを後回しにします。休・応援などで出勤しない日の前に、その人の鍵を出勤する他の人へあらかじめ渡し、「鍵を持っていない」NGが出ないようにします（この設定がONの間は鍵No.3も固定せず受け渡します。画面2の「ルール」でON/OFF・人数を変更できます）。</li>
                            <li><strong>金庫チェック連日NG</strong>：金庫チェックの連日担当のみNGとします。</li>
                            <li><strong>NGの自動修正</strong>：割り当て結果でNGが出たら「自動で修正する」を押すと、条件を追加して作り直した「3-2 割り当て結果」「3-3 …」を作成します（NGが無くなるまで押せます）。「手動で一気に直す」で修正を飛ばして手動修正画面へ進むこともできます。</li>
                        </ul>
                    </div>
                    <div id="rulesPanel" class="space-y-4"></div>
                </div>
            </div>
        </section>
 
        <!-- 画面3: 割り当て結果 -->
        <section id="tab3" class="tab-content hidden space-y-6">
            <div id="errorPanel"></div>
            <div id="fixPanel"></div>
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                    <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-1">対象期間・タイトル</div>
                    <div id="summaryPeriod" class="text-base font-bold text-slate-700">-</div>
                </div>
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                    <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-1">未設定・NG日数</div>
                    <div id="summaryWarnings" class="text-base font-bold text-emerald-600">0件</div>
                </div>
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                    <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-1">検証ステータス</div>
                    <div id="summaryStatus" class="text-base font-bold text-slate-400">-</div>
                </div>
            </div>
 
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <h3 class="font-bold text-slate-700 mb-3 flex items-center gap-2 text-sm sm:text-base">
                    <i class="fa-solid fa-chart-pie text-indigo-600"></i> 担当者別の担当回数集計
                </h3>
                <div class="overflow-x-auto">
                    <table class="min-w-full text-xs sm:text-sm text-center text-slate-600 divide-y divide-slate-200 border border-slate-200 rounded-lg">
                        <thead class="bg-slate-100 text-slate-700">
                            <tr>
                                <th class="px-3 py-2 font-semibold text-center whitespace-nowrap">氏名</th>
                                <th class="px-3 py-2 font-semibold text-center whitespace-nowrap">優先順位</th>
                                <th class="px-3 py-2 font-semibold text-center whitespace-nowrap">鍵開け</th>
                                <th class="px-3 py-2 font-semibold text-center whitespace-nowrap">鍵閉め</th>
                                <th class="px-3 py-2 font-semibold text-center whitespace-nowrap">金庫チェック</th>
                                <th class="px-3 py-2 font-semibold text-center whitespace-nowrap">合計</th>
                            </tr>
                        </thead>
                        <tbody id="statsTableBody" class="divide-y divide-slate-200 bg-white"></tbody>
                    </table>
                </div>
            </div>
 
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4 gap-2">
                    <h3 class="font-bold text-slate-700 flex items-center gap-2 text-sm sm:text-base">
                        <i class="fa-solid fa-calendar-check text-indigo-600"></i> <span id="resultVerLabel" class="px-2 py-0.5 rounded-md bg-indigo-600 text-white text-xs">3-1</span> 割り当て結果 ＆ 最終チェック (OK/NG)
                    </h3>
                    <button type="button" onclick="recalculate()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-3.5 py-1.5 rounded-lg text-xs sm:text-sm transition flex items-center gap-1.5 shadow-sm cursor-pointer">
                        <i class="fa-solid fa-rotate"></i> 再計算する
                    </button>
                </div>
                <div class="flex flex-wrap items-center gap-x-4 gap-y-2 mb-3 text-xs text-slate-600">
                    <span class="font-semibold">凡例：</span>
                    <span class="px-2 py-0.5 rounded-md bg-blue-200 text-blue-900 font-bold">超早</span>
                    <span class="px-2 py-0.5 rounded-md bg-sky-100 text-sky-900 font-bold">早（9時まで出勤の時刻表記も含む）</span>
                    <span class="px-2 py-0.5 rounded-md bg-amber-100 text-amber-900 font-bold">遅</span>
                    <span class="px-2 py-0.5 rounded-md bg-violet-100 text-violet-900 font-bold">フル（早・遅どちらも可）</span>
                    <span class="px-2 py-0.5 rounded-md text-slate-400 font-bold">休・応援など</span>
                    <span class="text-slate-300">｜</span>
                    <span class="inline-flex items-center gap-1.5"><span class="inline-block w-9 h-5 rounded-md bg-slate-100 border border-slate-300"></span>引渡者＝前日から鍵を持っている人</span>
                    <span class="text-slate-400 font-bold">→</span>
                    <span class="inline-flex items-center gap-1.5"><span class="inline-block w-9 h-5 rounded-md bg-indigo-100 border-2 border-indigo-500"></span>引受者＝その日に新しく受け取る人（太枠）</span>
                    <span class="inline-flex items-center gap-1"><span class="text-slate-400 font-bold">―</span>引受者が空欄＝同じ人が持ち続ける（受け渡しなし）</span>
                    </div>
                <div class="overflow-auto border border-slate-200 rounded-lg max-h-[72vh]">
                    <table class="compact-table w-full text-xs text-center text-slate-600 border-separate border-spacing-0">
                        <thead class="text-slate-700">
                            <tr>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-slate-50 left-0 z-30" rowspan="2">日付</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-slate-50 text-slate-700 border-l-2 border-slate-300" id="resultShiftTh" colspan="5">シフト状況</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-slate-50 border-l-2 border-slate-300" rowspan="2">鍵開け</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-slate-50" rowspan="2">鍵閉め</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-slate-50" rowspan="2">金庫チェック</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-red-50 text-slate-800 border-t-4 border-red-400 border-l-2 border-l-slate-300" colspan="2">鍵No.1 (赤)</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-amber-50 text-slate-800 border-t-4 border-amber-400 border-l-2 border-l-slate-300" colspan="2">鍵No.2 (黄)</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-blue-50 text-slate-800 border-t-4 border-blue-400 border-l-2 border-l-slate-300" colspan="2">鍵No.3 (青)</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-slate-50 border-l-2 border-slate-300" rowspan="2">最終</th>
                                <th class="sticky top-0 z-20 h-11 px-2 font-semibold whitespace-nowrap text-center bg-slate-50 min-w-[180px]" rowspan="2">詳細理由</th>
                            </tr>
                            <tr id="resultHeadRow2">
                                <th class="sticky top-[44px] z-20 px-2 py-1.5 text-[11px] whitespace-nowrap text-center bg-red-50 text-slate-600 border-l-2 border-slate-300 min-w-[4.5rem]">引渡者</th>
                                <th class="sticky top-[44px] z-20 px-2 py-1.5 text-[11px] whitespace-nowrap text-center bg-red-50 text-slate-600 min-w-[4.5rem]">引受者</th>
                                <th class="sticky top-[44px] z-20 px-2 py-1.5 text-[11px] whitespace-nowrap text-center bg-amber-50 text-slate-600 border-l-2 border-slate-300 min-w-[4.5rem]">引渡者</th>
                                <th class="sticky top-[44px] z-20 px-2 py-1.5 text-[11px] whitespace-nowrap text-center bg-amber-50 text-slate-600 min-w-[4.5rem]">引受者</th>
                                <th class="sticky top-[44px] z-20 px-2 py-1.5 text-[11px] whitespace-nowrap text-center bg-blue-50 text-slate-600 border-l-2 border-slate-300 min-w-[4.5rem]">引渡者</th>
                                <th class="sticky top-[44px] z-20 px-2 py-1.5 text-[11px] whitespace-nowrap text-center bg-blue-50 text-slate-600 min-w-[4.5rem]">引受者</th>
                            </tr>
                        </thead>
                        <tbody id="resultTableBody" class="divide-y divide-slate-200 bg-white"></tbody>
                    </table>
                </div>
            </div>
        </section>
 
        <!-- 画面3-比較: 版の比較 -->
        <section id="tabCmp" class="tab-content hidden space-y-6">
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <h2 class="text-base font-bold text-slate-700 mb-3 flex items-center gap-2">
                    <i class="fa-solid fa-table-columns text-indigo-600"></i> 版の比較（3-1・3-2 … を左右に並べて見比べる）
                </h2>
                <div id="cmpBody"></div>
            </div>
        </section>

        <!-- 画面4: 手動修正 -->
        <section id="tab4" class="tab-content hidden space-y-6">
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <h2 class="text-base font-bold text-slate-700 mb-2 flex items-center gap-2">
                    <i class="fa-solid fa-pen-to-square text-indigo-600"></i> シフトの手動修正・担当固定
                </h2>
                <p class="text-xs sm:text-sm text-slate-500 mb-4">
                    担当者や鍵1、2、3の保持者を変更できます（鍵が回らない場合は「店舗」も選べます）。「固定」にチェックを入れると再計算時に上書きされません。
                </p>
                <div id="editToolbar" class="flex flex-wrap items-center gap-3 mb-3 p-3 rounded-lg border border-slate-200 bg-slate-50 text-xs">
                    <label class="flex items-center gap-1.5 cursor-pointer font-bold text-slate-700"><input type="checkbox" id="editNgOnly" onchange="renderEditTable()"> NGの日だけ表示する</label>
                    <span id="editNgCount" class="text-slate-500"></span>
                    <span class="text-slate-400">※変更した日は自動で「固定」になります。直し終わったら下の「変更を反映して再計算」を押してください。</span>
                </div>
                <div class="overflow-auto border border-slate-200 rounded-lg max-h-[500px]">
                    <table class="compact-table w-full text-xs text-center text-slate-600 divide-y divide-slate-200">
                        <thead id="editTableHead" class="bg-slate-100 text-slate-700 sticky top-0 z-10"></thead>
                        <tbody id="editTableBody" class="divide-y divide-slate-200 bg-white"></tbody>
                    </table>
                </div>
                <div class="mt-4 flex flex-wrap justify-end gap-3">
                    <button type="button" onclick="reflectManualEdits(false)" class="bg-white border border-indigo-300 text-indigo-700 hover:bg-indigo-50 font-medium px-5 py-2 rounded-lg shadow-sm transition flex items-center gap-2 cursor-pointer text-sm">
                        <i class="fa-solid fa-check"></i> 変更を反映する（新しい「3-○ 割り当て結果」を作成）
                    </button>
                    <button type="button" onclick="applyAndGoExport()" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-5 py-2 rounded-lg shadow transition flex items-center gap-2 cursor-pointer text-sm">
                        <i class="fa-solid fa-file-excel"></i> 変更を反映してExcel出力へ進む<span class="ml-1 px-1.5 py-0.5 rounded bg-white/25 text-[11px]">（推奨）</span> <i class="fa-solid fa-arrow-right"></i>
                    </button>
                </div>
            </div>
        </section>
 
        <!-- 画面5: Excel出力 -->
        <section id="tab5" class="tab-content hidden space-y-6">
            <div id="exportStatus"></div>
            <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
                <!-- 出力1: システム標準形式 -->
                <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-200 text-center py-8">
                    <div class="w-14 h-14 bg-indigo-50 text-indigo-600 rounded-full flex items-center justify-center mx-auto mb-3 text-xl border border-indigo-200">
                        <i class="fa-solid fa-file-excel"></i>
                    </div>
                    <h2 class="text-base font-bold text-slate-700 mb-2">システム集計版 Excel出力</h2>
                    <p class="text-xs text-slate-500 mb-5">
                        シフト状況、担当者、鍵保持者、判定結果をまとめた標準レイアウトのExcelファイルを出力します。
                    </p>
                    <button type="button" onclick="exportToExcel()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold px-5 py-2 rounded-lg shadow transition inline-flex items-center gap-2 text-xs cursor-pointer">
                        <i class="fa-solid fa-download"></i> 標準形式でダウンロード (.xls)
                    </button>
                </div>
 
                <!-- 出力2: 添付Excel完全準拠形式 -->
                <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-200 text-center py-8">
                    <div class="w-14 h-14 bg-emerald-50 text-emerald-600 rounded-full flex items-center justify-center mx-auto mb-3 text-xl border border-emerald-200">
                        <i class="fa-solid fa-file-lines"></i>
                    </div>
                    <h2 class="text-base font-bold text-slate-700 mb-2">【添付ファイル完全準拠】鍵管理表出力</h2>
                    <p class="text-xs text-slate-500 mb-5">
                        「10月鍵開け2.xlsx」のタイトル・5/6/7行目の文言・フォント・セル結合・引渡者/引受者仕様（同値時空欄）に完全に準拠して出力します。
                    </p>
                    <button type="button" onclick="exportToAttachedExcelStyle()" class="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-5 py-2 rounded-lg shadow transition inline-flex items-center gap-2 text-xs cursor-pointer">
                        <i class="fa-solid fa-download"></i> 添付形式でダウンロード (.xls)
                    </button>
                </div>

                <!-- 出力3: 添付形式＋シフト状況（確認用） -->
                <div class="bg-white p-6 rounded-xl shadow-sm border-2 border-amber-300 text-center py-8 relative">
                    <span class="absolute top-3 right-3 text-[10px] font-bold px-2 py-0.5 rounded-full bg-amber-100 text-amber-800">確認作業におすすめ</span>
                    <div class="w-14 h-14 bg-amber-50 text-amber-600 rounded-full flex items-center justify-center mx-auto mb-3 text-xl border border-amber-200">
                        <i class="fa-solid fa-table-columns"></i>
                    </div>
                    <h2 class="text-base font-bold text-slate-700 mb-2">【添付ファイル完全準拠】鍵管理表＋シフト状況（確認用）</h2>
                    <p class="text-xs text-slate-500 mb-5">
                        添付形式の鍵管理表の<strong>右側</strong>に、その日の全員のシフト・鍵開け／鍵閉め／金庫チェック担当・判定（OK/NG）を並べて出力します。シフトの種類・鍵の色・NGの種類ごとに色分けし、一目で確認できます。
                    </p>
                    <button type="button" onclick="exportToCheckStyle()" class="bg-amber-500 hover:bg-amber-600 text-white font-semibold px-5 py-2 rounded-lg shadow transition inline-flex items-center gap-2 text-xs cursor-pointer">
                        <i class="fa-solid fa-download"></i> 確認用でダウンロード (.xls)
                    </button>
                </div>
            </div>
        </section>
 
    </main>
 
    <!-- フッター -->
    <footer class="bg-white border-t border-slate-200 mt-12 py-4 text-center text-xs text-slate-400">
        &copy; 2026 鍵当番・金庫チェック自動割当システム All Rights Reserved.
    </footer>
 
    <!-- JavaScript ロジック -->
    <script>
        let rawShiftData = [];
        let detectedExcelTitle = "";
        let staffList = [
            { name: "五家", role: "リーダー", aliases: ["五家夏希"], mondaySafe: false, priority: 1, openCount: 0, closeCount: 0, safeCount: 0 },
            { name: "爲貝", role: "リーダー", aliases: ["爲貝虹太"], mondaySafe: false, priority: 2, openCount: 0, closeCount: 0, safeCount: 0 },
            { name: "柿沼", role: "店長",     aliases: ["柿沼優哉"], mondaySafe: true,  priority: 3, openCount: 0, closeCount: 0, safeCount: 0 },
            { name: "村瀬", role: "リーダー", aliases: ["村瀨慎輔", "村瀬慎輔", "村瀨"], mondaySafe: false, priority: 4, openCount: 0, closeCount: 0, safeCount: 0 },
            { name: "楠美", role: "リーダー", aliases: ["楠美結月"], mondaySafe: false, priority: 5, openCount: 0, closeCount: 0, safeCount: 0 }
        ];
        let shiftResult = [];
        let fixedKeyHolder = "五家"; // 鍵No.3を持ち回す人（画面2で変更可）
 
        const STORE_NAME = "店舗"; // 鍵を店舗に置く場合の表記
        const UNSET = "担当者未設定";
 
        function switchTab(tabId) {
            try {
                document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
                document.querySelectorAll('.tab-btn').forEach(el => {
                    el.classList.remove('border-indigo-600', 'text-indigo-600');
                    el.classList.add('border-transparent', 'text-slate-500');
                });
                const targetEl = document.getElementById(tabId);
                if (targetEl) targetEl.classList.remove('hidden');
                const activeBtn = document.getElementById('btn-' + tabId);
                if (activeBtn) {
                    activeBtn.classList.remove('border-transparent', 'text-slate-500');
                    activeBtn.classList.add('border-indigo-600', 'text-indigo-600');
                }
 
                if (tabId === 'tab3') renderResult();
                if (tabId === 'tab4') renderEditTable();
            } catch (err) {
                console.error("Tab switch error:", err);
            }
        }
 
        function setLoadStatus(msg, success) {
            const el = document.getElementById('loadStatus');
            if (!el) return;
            el.textContent = msg;
            el.className = success
                ? "text-xs px-2.5 py-1 bg-emerald-50 text-emerald-700 rounded-full font-medium border border-emerald-200"
                : "text-xs px-2.5 py-1 bg-amber-50 text-amber-700 rounded-full font-medium border border-amber-200";
        }
 
        // ---------------------------------------------------------------------
        // Excel読込パーサー（年月の自動判定・日付行の柔軟検出）
        // ---------------------------------------------------------------------
        const DOW = ['日', '月', '火', '水', '木', '金', '土'];
        let detectedYear = null;
        let detectedMonth = null;
        let lastLoaded = null; // 直近に読み込んだシート（年月変更時の再解析用）

        function getSelectedYM() {
            const y = parseInt(document.getElementById('selYear').value, 10);
            const m = parseInt(document.getElementById('selMonth').value, 10);
            return { y: y || null, m: m || null };
        }

        // 年月が未選択ならメッセージを出して false を返す
        function requireYM() {
            const ym = getSelectedYM();
            if (!ym.y || !ym.m) {
                alert("先に「① 対象の年月」を選択してください。");
                return false;
            }
            return true;
        }

        function toHalfWidth(s) {
            return String(s).replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
        }

        function getFallbackYM() {
            const now = new Date();
            return {
                y: detectedYear || now.getFullYear(),
                m: detectedMonth || (now.getMonth() + 1)
            };
        }

        // タイトル行・ファイル名・シート名から年月を推定
        function guessContextYM(rows, fileName, sheetName) {
            let y = null, m = null;
            const re = /(令和|平成|R|H)?(\d{1,4})年(\d{1,2})月/;
            for (let r = 0; r < Math.min(rows.length, 8) && m === null; r++) {
                const row = rows[r];
                if (!row) continue;
                for (let c = 0; c < row.length; c++) {
                    if (row[c] === undefined || row[c] === null || row[c] instanceof Date) continue;
                    const s = toHalfWidth(String(row[c]).replace(/[\s　]/g, ""));
                    const mt = s.match(re);
                    if (mt) {
                        const n = parseInt(mt[2], 10);
                        const era = mt[1];
                        if (era === "平成" || era === "H") y = 1988 + n;
                        else if (era === "令和" || era === "R" || n < 100) y = 2018 + n;
                        else y = n;
                        m = parseInt(mt[3], 10);
                        break;
                    }
                }
            }
            if (m === null) {
                for (const s0 of [fileName, sheetName]) {
                    const mt = toHalfWidth(String(s0 || "")).match(/(\d{1,2})月/);
                    if (mt && parseInt(mt[1], 10) >= 1 && parseInt(mt[1], 10) <= 12) {
                        m = parseInt(mt[1], 10);
                        break;
                    }
                }
            }
            const now = new Date();
            return { y: y || now.getFullYear(), m: m || null };
        }

        // 1セルを日付として解釈（日付型 / シリアル値 / "10/5" / "5日" / 数字のみ）
        function parseDateCell(cell, ctx) {
            if (cell === undefined || cell === null || cell === "") return null;
            const ctxM = ctx.m || (new Date().getMonth() + 1);

            if (cell instanceof Date) {
                const dt = new Date(cell.getTime() + 12 * 3600 * 1000); // タイムゾーン誤差吸収
                return { m: dt.getMonth() + 1, d: dt.getDate(), dow: DOW[dt.getDay()], dayOnly: false };
            }

            if (typeof cell === "number") {
                if (cell >= 40000 && cell <= 60000) {
                    const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(cell) * 86400000);
                    return { m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: DOW[dt.getUTCDay()], dayOnly: false };
                }
                if (Number.isInteger(cell) && cell >= 1 && cell <= 31) {
                    return { m: ctxM, d: cell, dow: DOW[new Date(ctx.y, ctxM - 1, cell).getDay()], dayOnly: true };
                }
                return null;
            }

            const s = toHalfWidth(String(cell).trim());
            let mt = s.match(/^(\d{1,2})[\/月.\-](\d{1,2})日?$/);
            if (mt) {
                const mm = parseInt(mt[1], 10), dd = parseInt(mt[2], 10);
                if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
                    let yy = ctx.y;
                    if (ctx.m) {
                        if (mm < ctx.m - 6) yy += 1;
                        else if (mm > ctx.m + 6) yy -= 1;
                    }
                    return { m: mm, d: dd, dow: DOW[new Date(yy, mm - 1, dd).getDay()], dayOnly: false };
                }
            }
            mt = s.match(/^(\d{1,2})日?\s*[\(（]?[日月火水木金土]?[\)）]?$/);
            if (mt) {
                const dd = parseInt(mt[1], 10);
                if (dd >= 1 && dd <= 31) {
                    return { m: ctxM, d: dd, dow: DOW[new Date(ctx.y, ctxM - 1, dd).getDay()], dayOnly: true };
                }
            }
            return null;
        }

        // 日付行を検出（先頭15行から最も日付が多い行を採用）
        function detectDateRow(rows, ctx) {
            let best = { rowIdx: -1, cols: [] };
            for (let r = 0; r < Math.min(rows.length, 15); r++) {
                const row = rows[r];
                if (!row) continue;
                const found = [];
                for (let c = 0; c < row.length; c++) {
                    const p = parseDateCell(row[c], ctx);
                    if (p) found.push({ col: c, m: p.m, d: p.d, dow: p.dow, dayOnly: p.dayOnly });
                }
                if (found.length < 10) continue;
                // 数字のみの行は「ほぼ昇順」であることを要求（他の数値行の誤検出防止）
                if (found.every(f => f.dayOnly)) {
                    let inc = 0;
                    for (let i = 1; i < found.length; i++) if (found[i].d > found[i - 1].d) inc++;
                    if (inc < (found.length - 1) * 0.8) continue;
                }
                if (found.length > best.cols.length) best = { rowIdx: r, cols: found };
            }
            return best;
        }

        document.getElementById('shiftFile').addEventListener('change', function(ev) {
            const input = ev.target;
            const file = input.files[0];
            if (!file) return;
            if (!requireYM()) { input.value = ""; return; }
            const fileName = file.name;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                    // 全シートを調べ、日付行が最も良く取れるシートを採用
                    let best = null;
                    for (const name of workbook.SheetNames) {
                        const ws = workbook.Sheets[name];
                        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
                        if (!rows || rows.length === 0) continue;
                        const ctx = getSelectedYM();
                        const det = detectDateRow(rows, ctx);
                        if (!best || det.cols.length > best.det.cols.length) {
                            best = { name, rows, ctx, det };
                        }
                    }
                    if (!best) throw new Error("no sheet");

                    lastLoaded = { rows: best.rows, name: best.name };
                    parseExcelExactLayout(best.rows, best.name, getSelectedYM(), best.det);
                } catch (err) {
                    console.error(err);
                    alert("ファイルの解析に失敗しました。ファイル形式や内容を確認してください。");
                    resetLoadedData();
                } finally {
                    input.value = ""; // 同じファイルを再選択しても読み込めるように
                }
            };
            reader.readAsArrayBuffer(file);
        });

        function parseExcelExactLayout(rows, sheetName, ctx, det) {
            // 登録メンバー（氏名＋Excel上の表記）から照合表を作成。長い表記を優先
            const matchKeys = [];
            staffList.forEach(s => {
                [s.name, ...(s.aliases || [])].forEach(k => {
                    k = String(k || "").trim();
                    if (k) matchKeys.push({ key: k, canonical: s.name });
                });
            });
            matchKeys.sort((a, b) => b.key.length - a.key.length);

            let parsedItems = [];
            let dateRowIdx = det ? det.rowIdx : -1;
            let dateColumns = [];

            const selY = ctx.y, selM = ctx.m;
            const lastDay = new Date(selY, selM, 0).getDate();
            detectedYear = selY;
            detectedMonth = selM;
            let monthNote = "";

            if (det && det.cols.length > 0) {
                // 指定年月を正とし、シフト表の「日にち」だけを使う
                let prev = 0;
                const seen = new Set();
                for (const f of det.cols) {
                    if (f.d < prev) break;              // 翌月分に入ったら打ち切り
                    prev = f.d;
                    if (f.d > lastDay || seen.has(f.d)) continue;
                    seen.add(f.d);
                    dateColumns.push({ col: f.col, dateStr: `${selM}/${f.d}`, dayOfWeek: DOW[new Date(selY, selM - 1, f.d).getDay()] });
                }
                // ファイル内の日付の月と指定月が違う場合は注意表示
                const cnt = {};
                det.cols.filter(f => !f.dayOnly).forEach(f => { cnt[f.m] = (cnt[f.m] || 0) + 1; });
                const keys = Object.keys(cnt);
                if (keys.length > 0) {
                    const fm = parseInt(keys.sort((a, b) => cnt[b] - cnt[a])[0], 10);
                    if (fm !== selM) monthNote = `（※ファイル内の日付は${fm}月ですが、指定の${selM}月として処理）`;
                }
            }
            if (dateColumns.length === 0) {
                for (let d = 1; d <= lastDay; d++) {
                    dateColumns.push({ col: d + 2, dateStr: `${selM}/${d}`, dayOfWeek: DOW[new Date(selY, selM - 1, d).getDay()] });
                }
            }

            // タイトル決定（ファイル内のタイトルが実データの月と一致する場合のみ採用）
            let titleFound = null;
            for (let r = 0; r < Math.min(rows.length, 4) && !titleFound; r++) {
                const row = rows[r];
                if (!row) continue;
                for (let c = 0; c < row.length; c++) {
                    if (row[c] instanceof Date) continue;
                    const s = String(row[c] === undefined || row[c] === null ? "" : row[c]).trim();
                    if (s.includes("年") && s.includes("月")) { titleFound = s; break; }
                }
            }
            const genTitle = `令和　${detectedYear - 2018}年 ${detectedMonth}月度　鍵管理表予定表【3か月保管】`;
            if (titleFound) {
                const tm = toHalfWidth(titleFound.replace(/[\s　]/g, "")).match(/年(\d{1,2})月/);
                detectedExcelTitle = (tm && parseInt(tm[1], 10) === detectedMonth) ? titleFound : genTitle;
            } else {
                detectedExcelTitle = genTitle;
            }

            rows.forEach((row, rIdx) => {
                if (!row) return;
                if (dateRowIdx >= 0 ? rIdx <= dateRowIdx : rIdx <= 1) return;

                let matchedCanonical = null;
                for (let c = 0; c < row.length; c++) {
                    let cellStr = String(row[c] === undefined || row[c] === null ? "" : row[c]).trim();
                    if (!cellStr || cellStr.includes("ﾒﾝﾀｰ") || cellStr.includes("メンター") || cellStr.includes("部門") || cellStr.includes("店長・AV")) continue;

                    const exact = matchKeys.find(mk => cellStr === mk.key);
                    const partial = exact || matchKeys.find(mk => cellStr.includes(mk.key));
                    if (partial) matchedCanonical = partial.canonical;
                    if (matchedCanonical) break;
                }

                if (matchedCanonical) {
                    let alreadyAdded = parsedItems.some(item => item.name === matchedCanonical && item.date === dateColumns[0].dateStr);
                    if (!alreadyAdded) {
                        dateColumns.forEach(dc => {
                            let rawVal = row[dc.col];
                            let shiftVal = (rawVal !== undefined && rawVal !== null && String(rawVal).trim() !== "") ? String(rawVal).trim() : "公休";

                            parsedItems.push({
                                date: dc.dateStr,
                                dayOfWeek: dc.dayOfWeek,
                                name: matchedCanonical,
                                shiftType: shiftVal
                            });
                        });
                    }
                }
            });

            if (parsedItems.length === 0) {
                alert("登録メンバーの行がシフト表内で見つかりませんでした。メンバー名または「Excel上の表記」を確認してください。");
                resetLoadedData();
                return;
            }

            rawShiftData = parsedItems;
            shiftResult = [];
            renderPreview();
            setLoadStatus(`読込成功（${detectedYear}年${detectedMonth}月 / ${new Set(rawShiftData.map(i => i.date)).size}日分）${monthNote}`, true);
        }

        // 旧：サンプルデータ自動生成。Excel未読込のときは何も表示しない仕様のため無効化
        function generateDefaultSampleShift() {
            return false;
        }

        // 読込データを初期（未読込）状態に戻す
        function resetLoadedData() {
            rawShiftData = [];
            shiftResult = [];
            lastLoaded = null;
            detectedYear = null;
            detectedMonth = null;
            detectedExcelTitle = "";
            renderPreview();
            setLoadStatus("未読込", false);
        }

        function renderPreview() {
            renderStaffHeaders();
            const tbody = document.getElementById('previewBody');
            if (!tbody) return;
            tbody.innerHTML = "";
 
            const dateMap = {};
            rawShiftData.forEach(item => {
                if (!dateMap[item.date]) {
                    dateMap[item.date] = { date: item.date, dayOfWeek: item.dayOfWeek, shifts: {} };
                }
                dateMap[item.date].shifts[item.name] = item.shiftType;
            });
 
            Object.values(dateMap).forEach(dObj => {
                const tr = document.createElement('tr');
                let html = `<td class="px-3 py-2 font-medium whitespace-nowrap text-center">${dObj.date}</td>
                            <td class="px-3 py-2 whitespace-nowrap text-center">${dObj.dayOfWeek}</td>`;
                orderedStaff().forEach(st => {
                    let s = dObj.shifts[st.name] || "公休";
                    let badgeColor = ["公休", "休日", "有給", "応援", "欠勤", "方針"].includes(s) ? "bg-slate-100 text-slate-500" : "bg-indigo-50 text-indigo-700 font-semibold";
                    html += `<td class="px-3 py-2 text-center"><span class="px-2 py-0.5 rounded text-xs ${badgeColor}">${s}</span></td>`;
                });
                tr.innerHTML = html;
                tbody.appendChild(tr);
            });
        }
 
        // ---------------------------------------------------------------------
        // メンバー管理（追加・削除・氏名変更・役職・並び順）
        // ---------------------------------------------------------------------
        const MIN_STAFF = 3, MAX_STAFF = 12;

        function esc(v) {
            return String(v === undefined || v === null ? "" : v)
                .replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        }

        function orderedStaff() {
            return staffList.slice().sort((a, b) => a.priority - b.priority);
        }

        function normalizePriorities() {
            staffList.sort((a, b) => a.priority - b.priority);
            staffList.forEach((s, i) => { s.priority = i + 1; });
        }

        function getFixedHolder() {
            if (!staffList.some(s => s.name === fixedKeyHolder)) {
                const first = orderedStaff()[0];
                fixedKeyHolder = first ? first.name : "";
            }
            return fixedKeyHolder;
        }

        function validateStaffName(name, selfIndex) {
            if (!name) return "氏名を入力してください。";
            if (/["'<>&|\\`]/.test(name)) return "氏名に使えない文字（\" ' < > & | \\ `）が含まれています。";
            if (name === STORE_NAME || name === UNSET || name === "-") return `「${name}」は予約語のため氏名に使えません。`;
            if (staffList.some((s, i) => i !== selfIndex && s.name === name)) return `「${name}」は既に登録されています。`;
            return "";
        }

        function parseAliases(str) {
            return String(str || "").split(/[,、，\n]/).map(x => x.trim()).filter(Boolean);
        }

        function roleLabel(s) {
            return s && s.role ? `<span class="ml-1 text-[10px] text-slate-400">(${esc(s.role)})</span>` : "";
        }

        // 構成が変わったら結果を作り直す
        function onStaffChanged() {
            shiftResult = [];
            renderPriorityList();
            renderStaffHeaders();
            if (rawShiftData && rawShiftData.length > 0) renderPreview();
        }

        function reparseLoaded() {
            if (!lastLoaded) return false;
            const ctx = getSelectedYM();
            const det = detectDateRow(lastLoaded.rows, ctx);
            parseExcelExactLayout(lastLoaded.rows, lastLoaded.name, ctx, det);
            return true;
        }

        function addStaff() {
            const nameEl = document.getElementById('newStaffName');
            const roleEl = document.getElementById('newStaffRole');
            const aliasEl = document.getElementById('newStaffAlias');
            const name = nameEl.value.trim();
            if (staffList.length >= MAX_STAFF) { alert(`メンバーは最大${MAX_STAFF}名までです。`); return; }
            const err = validateStaffName(name, -1);
            if (err) { alert(err); return; }

            staffList.push({
                name: name, role: roleEl.value.trim(), aliases: parseAliases(aliasEl.value), mondaySafe: false,
                priority: staffList.length + 1, openCount: 0, closeCount: 0, safeCount: 0
            });
            nameEl.value = ""; roleEl.value = ""; aliasEl.value = "";
            shiftResult = [];

            // 読込済みシフト表から新メンバーの行を拾い直す
            let notFound = false;
            if (lastLoaded) {
                reparseLoaded();
                notFound = !rawShiftData.some(i => i.name === name);
            } else if (rawShiftData && rawShiftData.length > 0) {
                generateDefaultSampleShift();
            }
            onStaffChanged();
            if (notFound) {
                alert(`「${name}」の行がシフト表内で見つかりませんでした（全日「公休」扱い）。\nシフト表に書かれている名前を「Excel上の表記」に入力すると反映されます。`);
            }
        }

        function removeStaff(i) {
            const s = staffList[i];
            if (!s) return;
            if (staffList.length <= MIN_STAFF) { alert(`鍵が3つあるため、メンバーは最低${MIN_STAFF}名必要です。`); return; }
            if (!confirm(`「${s.name}」を削除します。よろしいですか？`)) return;
            staffList.splice(i, 1);
            normalizePriorities();
            rawShiftData = (rawShiftData || []).filter(it => it.name !== s.name);
            onStaffChanged();
        }

        function renameStaff(i, value) {
            const s = staffList[i];
            if (!s) return;
            const name = String(value || "").trim();
            if (name === s.name) return;
            const err = validateStaffName(name, i);
            if (err) { alert(err); renderPriorityList(); return; }
            const old = s.name;
            s.name = name;
            s.aliases = [];   // 別人への入れ替えでも旧名で誤マッチしないよう表記はクリア
            if (fixedKeyHolder === old) fixedKeyHolder = name;
            (rawShiftData || []).forEach(it => { if (it.name === old) it.name = name; });
            onStaffChanged();
        }

        function setRole(i, value) { if (staffList[i]) staffList[i].role = String(value || "").trim(); }

        function setAliases(i, value) {
            if (!staffList[i]) return;
            staffList[i].aliases = parseAliases(value);
            if (lastLoaded) { reparseLoaded(); onStaffChanged(); }
        }

        function setMondaySafe(i, checked) {
            if (!staffList[i]) return;
            staffList[i].mondaySafe = !!checked;
            shiftResult = [];
        }

        function setFixedHolder(value) {
            fixedKeyHolder = value;
            onStaffChanged();
        }

        function renderPriorityList() {
            const container = document.getElementById('priorityList');
            if (!container) return;
            container.innerHTML = "";
            normalizePriorities();
            const fixedName = getFixedHolder();

            const inputCls = "border border-slate-300 rounded-md px-2 py-1 text-xs bg-white focus:outline-none focus:ring-2 focus:ring-indigo-300";
            staffList.forEach((staff, index) => {
                const div = document.createElement('div');
                div.className = "p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2";
                div.innerHTML = `
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <div class="flex items-center gap-2">
                            <span class="w-6 h-6 rounded-full bg-indigo-100 text-indigo-700 flex items-center justify-center text-xs font-bold">${staff.priority}</span>
                            <input type="text" value="${esc(staff.name)}" title="氏名" onchange="renameStaff(${index}, this.value)" class="w-24 font-bold ${inputCls}">
                            <input type="text" list="roleOptions" value="${esc(staff.role)}" placeholder="役職" title="役職" oninput="setRole(${index}, this.value)" class="w-24 ${inputCls}">
                        </div>
                        <div class="flex items-center gap-1">
                            <button type="button" onclick="movePriority(${index}, -1)" class="p-1.5 hover:bg-slate-200 rounded text-slate-500 cursor-pointer" ${index === 0 ? 'disabled style="opacity:0.3"' : ''}><i class="fa-solid fa-arrow-up"></i></button>
                            <button type="button" onclick="movePriority(${index}, 1)" class="p-1.5 hover:bg-slate-200 rounded text-slate-500 cursor-pointer" ${index === staffList.length - 1 ? 'disabled style="opacity:0.3"' : ''}><i class="fa-solid fa-arrow-down"></i></button>
                            <button type="button" onclick="removeStaff(${index})" title="削除" class="p-1.5 hover:bg-red-100 rounded text-red-500 cursor-pointer"><i class="fa-solid fa-trash"></i></button>
                        </div>
                    </div>
                    <div class="flex flex-wrap items-center gap-x-3 gap-y-1 text-xs text-slate-500">
                        <label class="flex items-center gap-1 flex-1 min-w-[12rem]">Excel上の表記
                            <input type="text" value="${esc((staff.aliases || []).join(', '))}" placeholder="例: 五家夏希（複数はカンマ区切り）" onchange="setAliases(${index}, this.value)" class="flex-1 ${inputCls}">
                        </label>
                        <label class="flex items-center gap-1 cursor-pointer"><input type="checkbox" ${staff.mondaySafe ? 'checked' : ''} onchange="setMondaySafe(${index}, this.checked)"> 月曜の金庫チェック優先</label>
                    </div>
                `;
                container.appendChild(div);
            });

            const box = document.getElementById('fixedHolderBox');
            if (box) {
                box.innerHTML = `
                    <label class="text-xs font-bold text-slate-700 flex flex-wrap items-center gap-2">
                        <i class="fa-solid fa-key text-blue-500"></i> 鍵No.3（青）を持ち回す人
                        <select onchange="setFixedHolder(this.value)" class="${inputCls}">
                            ${orderedStaff().map(s => `<option value="${esc(s.name)}" ${s.name === fixedName ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}
                        </select>
                    </label>`;
            }
            document.querySelectorAll('.fixedHolderName').forEach(el => { el.textContent = fixedName; });
            document.querySelectorAll('.deferTopNames').forEach(el => { el.textContent = orderedStaff().slice(0, Math.max(0, Math.min(Number(RULES.deferTopN) || 0, staffList.length - 3))).map(x => x.name).join('・') || '（なし）'; });
        }

        // プレビュー表・結果表のメンバー列ヘッダを現在のメンバーで作り直す
        function renderStaffHeaders() {
            const names = orderedStaff().map(s => s.name);

            const prev = document.getElementById('previewHeadRow');
            if (prev) {
                while (prev.children.length > 2) prev.removeChild(prev.lastChild);
                prev.insertAdjacentHTML('beforeend', names.map(n =>
                    `<th class="px-3 py-2.5 whitespace-nowrap text-center"><span class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 border border-slate-300">${esc(n)}</span></th>`
                ).join(''));
            }

            const row2 = document.getElementById('resultHeadRow2');
            if (row2) {
                row2.querySelectorAll('.staff-th').forEach(el => el.remove());
                row2.insertAdjacentHTML('afterbegin', names.map((n, i) =>
                    `<th class="staff-th sticky top-[44px] z-20 px-2 py-1.5 text-[11px] whitespace-nowrap text-center bg-slate-50 text-slate-700${i === 0 ? ' border-l-2 border-slate-300' : ''}"><span class="px-2 py-0.5 rounded-md font-bold bg-white text-slate-800 border border-slate-300">${esc(n)}</span></th>`
                ).join(''));
            }
            const sh = document.getElementById('resultShiftTh');
            if (sh) sh.setAttribute('colspan', String(names.length));
        }

        function movePriority(index, direction) {
            const targetIndex = index + direction;
            if (targetIndex < 0 || targetIndex >= staffList.length) return;

            const temp = staffList[index].priority;
            staffList[index].priority = staffList[targetIndex].priority;
            staffList[targetIndex].priority = temp;

            shiftResult = [];
            staffList.sort((a, b) => a.priority - b.priority);
            renderPriorityList();
            renderStaffHeaders();
            if (rawShiftData && rawShiftData.length > 0) renderPreview();
        }

        // -------------------------------------------------------------------------
        // シフト判定関数
        // -------------------------------------------------------------------------
        function isWorkingShift(shiftStr) {
            let s = String(shiftStr).trim();
            const absentKeywords = ["公休", "休日", "有給", "欠勤", "方針", "グランド", "応援", "店長会議"];
            if (absentKeywords.some(k => s.includes(k))) return false;
            if (s === "" || s === "nan" || s === "undefined" || s === "null") return false;
            return true;
        }
 
        // 「9 1830」「9:00-18:30」のような時刻表記の開始時刻（時）を返す。時刻表記でなければ null
        function shiftStartHour(shiftStr) {
            const t = String(shiftStr).trim().replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
            const m = t.match(/^(\d+)/);
            if (!m) return null;
            const tok = m[1];
            const h = tok.length >= 3 ? Math.floor(parseInt(tok, 10) / 100) : parseInt(tok, 10);
            return (h >= 0 && h <= 24) ? h : null;
        }
 
        // 9時までに出勤する時刻表記は「早」として扱う
        const EARLY_START_HOUR = 9;
        function isEarlyByTime(shiftStr) {
            const h = shiftStartHour(shiftStr);
            return h !== null && h <= EARLY_START_HOUR;
        }
 
        // 「フル」は早番・遅番どちらにも対応可能
        function isFullShift(shiftStr) {
            const s = String(shiftStr).trim().replace(/ﾌﾙ/g, "フル");
            return s.includes("フル");
        }

        // 鍵開けできるのは「超早」「早」「フル」の人（または9時までに出勤する時刻表記の人）
        function canOpenShift(shiftStr) {
            let s = String(shiftStr).trim();
            if (!isWorkingShift(s)) return false;
            return s.includes("早") || isFullShift(s) || isEarlyByTime(s);
        }
 
        // 鍵閉め・金庫チェックができるのは「遅番」「フル」の人
        function isLateShift(shiftStr) {
            let s = String(shiftStr).trim();
            if (!isWorkingShift(s)) return false;
            if (isFullShift(s)) return true;
            if (s.includes("早")) return false;
            return s.includes("遅") || s.includes("L") || s.includes("Ｌ");
        }
 
        // 鍵保持者の重複チェック（店舗置きは重複OK）
        function keyDupMessage(kh) {
            if (!kh) return "";
            const vals = [kh["鍵No.1"], kh["鍵No.2"], kh["鍵No.3"]].filter(v => v && v !== STORE_NAME);
            return (new Set(vals).size !== vals.length) ? "鍵の保持者が重複しています（3つの鍵は別々の人が保持）" : "";
        }
 
        // 表示用の最終ステータス（鍵重複も加味）
        function getDisplayStatus(item) {
            if (!item) return { status: "OK", note: "" };
            const dup = keyDupMessage(item.keyHolders);
            if (dup) {
                const baseNote = (item.checkStatus === "NG") ? item.note : "";
                return { status: "NG", note: baseNote ? `${baseNote} / ${dup}` : `割当不可: ${dup}` };
            }
            return { status: item.checkStatus, note: item.note };
        }
 
        // -------------------------------------------------------------------------
        // 自動割り当てロジック
        // -------------------------------------------------------------------------
        function proceedToGenerate() {
            try {
                if (!rawShiftData || rawShiftData.length === 0) {
                    alert("先にExcelファイルを読み込んでください。");
                    return;
                }
                generateShift();
                if (!shiftResult || shiftResult.length === 0) return;
                switchTab('tab3');
            } catch (err) {
                console.error("Generate error:", err);
                alert("自動作成中にエラーが発生しました。データを再確認してください。");
            }
        }
 
        function recalculate() {
            try {
                if (!rawShiftData || rawShiftData.length === 0) {
                    alert("先にExcelファイルを読み込んでください。");
                    return;
                }
                generateShift();
                renderResult();
                renderEditTable();
                alert("シフトを再計算しました。");
            } catch (err) {
                console.error("Recalculate error:", err);
                alert("再計算中にエラーが発生しました。");
            }
        }
 
        // -------------------------------------------------------------------------
        // 鍵の持ち回しプラン（月全体を見て「店舗置き」を最小にする）
        //  状態 = 鍵No.1・No.2の保持者（鍵No.3は五家が固定）。前日→当日の受け渡しは
        //  「両方が出勤している」か「店舗を経由して出勤者が受け取る」場合のみ可能。
        //  休・応援・グランド等で出勤しない人は、鍵を持ったまま次の出勤日まで持ち帰る。
        // -------------------------------------------------------------------------
        function planKeyHolders(dates, dateMap, lockedMap, defaultStart) {
            if (RULES.deferTop) return planKeyHoldersDefer(dates, dateMap, lockedMap, defaultStart);   // 上位者の鍵は先回りで他の人へ
            const fixedName = getFixedHolder();
            const order = staffList.slice().sort((a, b) => a.priority - b.priority).map(s => s.name);
            const vals = [...order, STORE_NAME];
            const states = [];
            vals.forEach(a => vals.forEach(b => { if (a === b && a !== STORE_NAME) return; states.push([a, b]); }));
            const keyOf = s => s[0] + "|" + s[1];
 
            const kd0 = {};
            order.forEach(n => { kd0[n] = 0; });
            let cur = new Map();
            states.forEach(s => {
                if (s.includes(STORE_NAME) || s.includes(fixedName)) return;
                const dev = (s[0] === defaultStart[0] ? 0 : 0.5) + (s[1] === defaultStart[1] ? 0 : 0.5);
                cur.set(keyOf(s), { cost: dev, s: s, from: null, kd: kd0, closer: UNSET });
            });
            const startLayer = cur;
            const layers = [];
 
            dates.forEach(dateKey => {
                const st = (dateMap[dateKey] && dateMap[dateKey].staffs) ? dateMap[dateKey].staffs : {};
                const shiftOf = n => st[n] ? String(st[n]).trim() : "公休";
                const W = new Set(order.filter(n => isWorkingShift(shiftOf(n))));
                const E = order.filter(n => canOpenShift(shiftOf(n)));
                const L = order.filter(n => isLateShift(shiftOf(n)));
                const locked = lockedMap[dateKey];
                const present = x => x === STORE_NAME || W.has(x);
                const lateOnly = x => x !== STORE_NAME && isLateShift(shiftOf(x)) && !isFullShift(shiftOf(x)) && !canOpenShift(shiftOf(x));
                const earlyOnly = x => x !== STORE_NAME && canOpenShift(shiftOf(x)) && !isFullShift(shiftOf(x));
                const cands = locked ? [[locked.keyHolders["鍵No.1"], locked.keyHolders["鍵No.2"]]] : states;
                const next = new Map();
 
                for (const pv of cur.values()) {
                    const p = pv.s;
                    // 鍵開け：超早・早の誰かが前日から鍵を持っているか、店舗に鍵がある必要がある
                    const holders = new Set([...p, fixedName]);
                    const openPenalty = (!locked && RULES.openNeedsKey && E.length > 0 && !(RULES.storeOpen && p.includes(STORE_NAME)) && !E.some(e => holders.has(e))) ? (RULES.penOpen || 1000) : 0;
                    // 鍵開け優先者（既定：五家）が早番の日は、その人が前日から鍵を持っているようにする
                    const prefPen = (!locked && RULES.openPreferOn && E.includes(RULES.openPrefer) && !holders.has(RULES.openPrefer) && !(RULES.storeOpen && p.includes(STORE_NAME))) ? (Number(RULES.openPreferWeight) || 60) : 0;
 
                    for (const s of cands) {
                        let cost = pv.cost + openPenalty + prefPen;
                        let closer = UNSET;
                        let bad = false;
                        for (let k = 0; k < 2; k++) {
                            const a = p[k], b = s[k];
                            if (a !== b && (!(present(a) && present(b)) || (RULES.noLateToEarly && lateOnly(a) && earlyOnly(b)))) bad = true;   // 非出勤者は受け渡し不可
                            if (a !== b) cost += 2;                                    // 受け渡しは少ないほど良い
                        }
                        if (locked) {
                            if (bad) cost += 50;   // 手動固定日は強制（矛盾していても受け入れる）
                        } else {
                            if (bad) continue;
                            if (s[0] === s[1] && s[0] !== STORE_NAME) continue;        // 同じ人が2本持たない
                            s.forEach(x => {
                                if (x === STORE_NAME) cost += RULES.storeAllow ? 100 : 100000;                    // 店舗置きは最後の手段
                                else if (x === fixedName) cost += 500;               // 固定担当は鍵No.3
                            });
                            // 鍵閉め担当は退勤時に鍵が必要（優先順位が高い遅番の人を優先）
                            if (L.length > 0) {
                                let best = null;
                                L.forEach((c, idx) => {
                                    const access = (c === fixedName) || p.includes(c) || s.includes(c);
                                    const cc = (access ? 0 : (RULES.penClose || 300)) + idx * 3;
                                    if (!best || cc < best.cc) best = { c: c, cc: cc };
                                });
                                closer = best.c;
                                cost += best.cc;
                            }
                        }
                        // 鍵を持つ日数の偏りを少し抑える
                        const kd = Object.assign({}, pv.kd);
                        s.forEach(x => {
                            if (x !== STORE_NAME && kd[x] !== undefined) { cost += 0.01 * kd[x]; kd[x]++; }
                        });
                        const key = keyOf(s);
                        const ex = next.get(key);
                        if (!ex || cost < ex.cost) next.set(key, { cost: cost, s: s, from: keyOf(p), kd: kd, closer: closer });
                    }
                }
                layers.push(next);
                cur = next;
            });
 
            // 最小コストの経路をたどる
            let bestNode = null;
            for (const v of cur.values()) { if (!bestNode || v.cost < bestNode.cost) bestNode = v; }
            const days = new Array(dates.length);
            let node = bestNode;
            for (let i = dates.length - 1; i >= 0; i--) {
                days[i] = { slots: node.s, closer: node.closer };
                node = (i === 0) ? startLayer.get(node.from) : layers[i - 1].get(node.from);
            }
            return { start: node ? node.s : defaultStart, days: days };
        }
 
        // -------------------------------------------------------------------------
        // 鍵の持ち回しプラン【上位者後回し版】
        //  優先順位の上位N名（既定：1位五家・2位爲貝）には鍵を持たせるのを後回しにし、
        //  その人たちが休・応援などで出勤しなくなる前に、鍵を出勤する他の人へあらかじめ渡しておく。
        //  こうして「前日の夜までに鍵を持っていない」「退勤時に鍵を持っていない」というNGを出さない。
        //  鍵No.3も固定せず、3つの鍵すべてを動かせる（状態 = 鍵No.1・No.2・No.3の保持者）。
        //  ※高速化のため、メンバーを番号に置き換えて計算している。
        // -------------------------------------------------------------------------
        function planKeyHoldersDefer(dates, dateMap, lockedMap, defaultStart) {
            const fixedName = getFixedHolder();
            const order = staffList.slice().sort((a, b) => a.priority - b.priority).map(s => s.name);
            const n = order.length, STORE = n, B = n + 1, S = B * B * B;
            const idxOf = {};
            order.forEach((nm, i) => { idxOf[nm] = i; });
            idxOf[STORE_NAME] = STORE;
            const nameOf = i => (i === STORE ? STORE_NAME : order[i]);
            const topN = Math.max(0, Math.min(Number(RULES.deferTopN) || 0, n - 3));   // 鍵3本を持てる人は最低3名残す
            const DEFER_COST = 10;      // 上位者が鍵を持つ1本・1日あたりの後回しコスト（店舗置き100より軽く、受け渡し2より重い）
            const BEAM = Math.max(50, Number(RULES.beam) || 150);           // 探索幅（毎日、コストの低い上位の状態だけ残して速度を保つ）
            const storeCost = RULES.storeAllow ? 100 : 100000;
            const slotCost = new Float64Array(B);
            for (let i = 0; i < n; i++) slotCost[i] = i < topN ? DEFER_COST : 0;
            slotCost[STORE] = storeCost;
            const enc = (a, b, c) => (a * B + b) * B + c;
            const prefIdx = (RULES.openPreferOn && idxOf[RULES.openPrefer] !== undefined && RULES.openPrefer !== STORE_NAME) ? idxOf[RULES.openPrefer] : -1;
            const validState = new Uint8Array(S);
            const A = new Int16Array(S), Bk = new Int16Array(S), C = new Int16Array(S);
            for (let a = 0; a < B; a++) for (let b = 0; b < B; b++) for (let c = 0; c < B; c++) {
                const id = enc(a, b, c);
                A[id] = a; Bk[id] = b; C[id] = c;
                validState[id] = !((a !== STORE && (a === b || a === c)) || (b !== STORE && b === c)) ? 1 : 0;
            }
            const holdCost = new Float64Array(S);
            for (let id = 0; id < S; id++) if (validState[id]) holdCost[id] = slotCost[A[id]] + slotCost[Bk[id]] + slotCost[C[id]];

            // 初期状態（月初は前日の保持者が分からないので、既定の保持者に近い並びを好む）
            const ref = [idxOf[defaultStart[0]], idxOf[defaultStart[1]], idxOf[fixedName]];
            let cost = new Float64Array(S).fill(Infinity);
            let kdArr = new Array(S).fill(null);
            for (let id = 0; id < S; id++) {
                if (!validState[id] || A[id] === STORE || Bk[id] === STORE || C[id] === STORE) continue;
                cost[id] = (A[id] === ref[0] ? 0 : 0.5) + (Bk[id] === ref[1] ? 0 : 0.5) + (C[id] === ref[2] ? 0 : 0.5);
                kdArr[id] = new Int16Array(n);
            }
            const startCost = cost.slice();
            const parents = [], closers = [];

            dates.forEach(dateKey => {
                const st = (dateMap[dateKey] && dateMap[dateKey].staffs) ? dateMap[dateKey].staffs : {};
                const shiftOf = nm => st[nm] ? String(st[nm]).trim() : "公休";
                const present = new Uint8Array(B), lateOnly = new Uint8Array(B), earlyOnly = new Uint8Array(B);
                const E = [], L = [];
                present[STORE] = 1;
                for (let i = 0; i < n; i++) {
                    const sh = shiftOf(order[i]);
                    if (isWorkingShift(sh)) present[i] = 1;
                    const canO = canOpenShift(sh), isL = isLateShift(sh), full = isFullShift(sh);
                    if (canO) E.push(i);
                    if (isL) L.push(i);
                    lateOnly[i] = (present[i] && isL && !full && !canO) ? 1 : 0;
                    earlyOnly[i] = (present[i] && canO && !full) ? 1 : 0;
                }
                // x から v へ受け渡せるか（両方が出勤している／遅番→早番は不可）
                const canPass = (x, v) => x !== v && present[x] && present[v] && !(RULES.noLateToEarly && lateOnly[x] && earlyOnly[v]);
                const moves = [];
                for (let x = 0; x < B; x++) {
                    const l = [x];
                    for (let v = 0; v < B; v++) if (canPass(x, v)) l.push(v);
                    moves.push(l);
                }
                // 手動固定日は、その保持者で強制
                let lockedId = -1;
                const lk = lockedMap[dateKey];
                if (lk) {
                    const la = idxOf[lk.keyHolders["鍵No.1"]], lb = idxOf[lk.keyHolders["鍵No.2"]], lc = idxOf[lk.keyHolders["鍵No.3"]];
                    if (la !== undefined && lb !== undefined && lc !== undefined) lockedId = enc(la, lb, lc);
                }

                const nCost = new Float64Array(S).fill(Infinity);
                const nPar = new Int32Array(S).fill(-1);
                const nCloser = new Int16Array(S).fill(-1);
                const nKd = new Array(S).fill(null);

                for (let pid = 0; pid < S; pid++) {
                    const pc0 = cost[pid];
                    if (pc0 === Infinity) continue;
                    const pa = A[pid], pb = Bk[pid], pcc = C[pid];
                    // 鍵開け：超早・早・フルの誰かが前日から鍵を持っているか、店舗に鍵がある必要がある
                    let openPenalty = 0;
                    if (lockedId < 0 && RULES.openNeedsKey && E.length > 0 && !(RULES.storeOpen && (pa === STORE || pb === STORE || pcc === STORE))) {
                        let ok = false;
                        for (let q = 0; q < E.length; q++) { const e = E[q]; if (e === pa || e === pb || e === pcc) { ok = true; break; } }
                        if (!ok) openPenalty = (RULES.penOpen || 1000);
                    }
                    if (lockedId < 0 && prefIdx >= 0 && E.includes(prefIdx) && pa !== prefIdx && pb !== prefIdx && pcc !== prefIdx && !(RULES.storeOpen && (pa === STORE || pb === STORE || pcc === STORE))) openPenalty += (Number(RULES.openPreferWeight) || 60);   // 鍵開け優先者は前日から鍵を持つ
                    const baseCost = pc0 + openPenalty;
                    const kdP = kdArr[pid];

                    const la = lockedId >= 0 ? [A[lockedId]] : moves[pa];
                    const lb = lockedId >= 0 ? [Bk[lockedId]] : moves[pb];
                    const lc = lockedId >= 0 ? [C[lockedId]] : moves[pcc];
                    for (let i1 = 0; i1 < la.length; i1++) {
                        const a = la[i1];
                        for (let i2 = 0; i2 < lb.length; i2++) {
                            const b = lb[i2];
                            if (b !== STORE && b === a) continue;
                            for (let i3 = 0; i3 < lc.length; i3++) {
                                const c = lc[i3];
                                if (c !== STORE && (c === a || c === b)) continue;
                                const sid = enc(a, b, c);
                                let cc = baseCost;
                                if (a !== pa) cc += 2;                                  // 受け渡しは少ないほど良い
                                if (b !== pb) cc += 2;
                                if (c !== pcc) cc += 2;
                                let closer = -1;
                                if (lockedId >= 0) {
                                    // 手動固定日は強制（矛盾していても受け入れる）
                                    if ((a !== pa && !canPass(pa, a)) || (b !== pb && !canPass(pb, b)) || (c !== pcc && !canPass(pcc, c))) cc += 50;
                                } else {
                                    cc += holdCost[sid];                                // 店舗置きは最後の手段・上位者に持たせるのは後回し
                                    // 鍵閉め担当は退勤時に鍵が必要（優先順位が高い遅番の人を優先）
                                    if (L.length > 0) {
                                        let found = false;
                                        for (let q = 0; q < L.length; q++) {
                                            const l = L[q];
                                            if (l === a || l === b || l === c || l === pa || l === pb || l === pcc) { cc += q * 3; closer = l; found = true; break; }
                                        }
                                        if (!found) { cc += (RULES.penClose || 300); closer = L[0]; }
                                    }
                                }
                                // 鍵を持つ日数の偏りを少し抑える
                                if (kdP) {
                                    if (a !== STORE) cc += 0.01 * kdP[a];
                                    if (b !== STORE) cc += 0.01 * kdP[b];
                                    if (c !== STORE) cc += 0.01 * kdP[c];
                                }
                                if (cc < nCost[sid]) {
                                    nCost[sid] = cc; nPar[sid] = pid; nCloser[sid] = closer;
                                    nKd[sid] = kdP;     // 実際に使われる状態だけ後で数え直す
                                }
                            }
                        }
                    }
                }
                // 探索幅を絞る
                const fin = [];
                for (let id = 0; id < S; id++) if (nCost[id] !== Infinity) fin.push(nCost[id]);
                if (fin.length > BEAM) {
                    fin.sort((x, y) => x - y);
                    const th = fin[BEAM - 1];
                    let kept = 0;
                    for (let id = 0; id < S; id++) {
                        if (nCost[id] === Infinity) continue;
                        if (nCost[id] > th || kept >= BEAM) { nCost[id] = Infinity; nPar[id] = -1; } else kept++;
                    }
                }
                // 鍵を持った日数を更新
                for (let id = 0; id < S; id++) {
                    if (nCost[id] === Infinity) { nKd[id] = null; continue; }
                    const k = Int16Array.from(nKd[id]);
                    if (A[id] !== STORE) k[A[id]]++;
                    if (Bk[id] !== STORE) k[Bk[id]]++;
                    if (C[id] !== STORE) k[C[id]]++;
                    nKd[id] = k;
                }
                parents.push(nPar); closers.push(nCloser);
                cost = nCost; kdArr = nKd;
            });

            // 最小コストの経路をたどる
            let best = -1;
            for (let id = 0; id < S; id++) if (cost[id] !== Infinity && (best < 0 || cost[id] < cost[best])) best = id;
            const slotsOf = id => [nameOf(A[id]), nameOf(Bk[id]), nameOf(C[id])];
            const days = new Array(dates.length);
            let node = best;
            for (let i = dates.length - 1; i >= 0; i--) {
                const cl = closers[i][node];
                days[i] = { slots: slotsOf(node), closer: cl >= 0 ? order[cl] : UNSET };
                node = parents[i][node];
            }
            return { start: node >= 0 ? slotsOf(node) : [defaultStart[0], defaultStart[1], fixedName], days: days };
        }

        function generateShift() {
            const dateMap = {};
            if (!rawShiftData || rawShiftData.length === 0) {
                if (!generateDefaultSampleShift()) { shiftResult = []; return; }
            }
            rawShiftData.forEach(item => {
                if (!item || !item.date) return;
                if (!dateMap[item.date]) {
                    dateMap[item.date] = { date: item.date, dayOfWeek: item.dayOfWeek || '木', staffs: {} };
                }
                if (item.name) {
                    dateMap[item.date].staffs[item.name] = item.shiftType || '公休';
                }
            });
 
            const dates = Object.keys(dateMap);
            if (dates.length === 0) {
                if (!generateDefaultSampleShift()) { shiftResult = []; return; }
                return generateShift();
            }
 
            let currentResultMap = {};
            if (shiftResult && shiftResult.length > 0) {
                shiftResult.forEach(r => {
                    if (r && r.locked && r.date) currentResultMap[r.date] = r;
                });
            }
 
            const fixedName = getFixedHolder();
            staffList.forEach(s => { s.openCount = 0; s.closeCount = 0; s.safeCount = 0; });
 
            // 鍵を持った日数（鍵を回す際の公平性に使用）
            const keyDays = {};
            staffList.forEach(s => { keyDays[s.name] = 0; });
            const priorityOf = (n) => { const s = staffList.find(x => x.name === n); return s ? s.priority : 99; };
 
            let newShiftResult = [];
            let lastSafeStaff = null;
            const safeHist = [];
            // 初期の鍵保持者（五家は鍵No.3を持ち回し）
            // 月初は「前日の保持者」が分からないため、初日の出勤者から決める
            //  （鍵開けできる超早・早の人を最優先、次に出勤者、最後に非番の人）
            const defaultStart = (() => {
                const d0 = (dateMap[dates[0]] && dateMap[dates[0]].staffs) ? dateMap[dates[0]].staffs : {};
                const order = staffList.slice().sort((a, b) => a.priority - b.priority).map(s => s.name);
                const working = order.filter(n => isWorkingShift(d0[n] || "公休"));
                const early = working.filter(n => canOpenShift(d0[n] || "公休"));
                const ordered = [];
                [...early, ...working, ...order].forEach(n => { if (!ordered.includes(n)) ordered.push(n); });
                const others = ordered.filter(n => n !== fixedName);   // 固定担当は鍵No.3を持ち回し
                return [others[0], others[1]];
            })();
            // 月全体を見渡して鍵の持ち回しを決める（店舗置きは「どうしても回らない時」だけ）
            const keyPlan = planKeyHolders(dates, dateMap, currentResultMap, defaultStart);
            let prevHolders = { "鍵No.1": keyPlan.start[0], "鍵No.2": keyPlan.start[1], "鍵No.3": keyPlan.start[2] || fixedName };
 
            dates.forEach((dateKey, index) => {
                const dayData = dateMap[dateKey] || { date: dateKey, dayOfWeek: '木', staffs: {} };
 
                if (currentResultMap[dateKey]) {
                    const locked = currentResultMap[dateKey];
                    newShiftResult.push(locked);
                    const openS = staffList.find(s => s.name === locked.openRaw);
                    const closeS = staffList.find(s => s.name === locked.closeRaw);
                    const safeS = staffList.find(s => s.name === locked.safeRaw);
                    if (openS) openS.openCount++;
                    if (closeS) closeS.closeCount++;
                    if (safeS && locked.safeRaw !== UNSET && locked.safeRaw !== "-") safeS.safeCount++;
                    if (locked.safeRaw !== UNSET && locked.safeRaw !== "-") { lastSafeStaff = locked.safeRaw; safeHist.push(locked.safeRaw); }
                    locked.prevKeyHolders = { ...prevHolders };
                    prevHolders = { ...locked.keyHolders };
                    return;
                }
 
                let availableStaffs = [];
                staffList.forEach(st => {
                    let name = st.name;
                    let shift = (dayData.staffs && dayData.staffs[name]) ? dayData.staffs[name] : "公休";
                    if (isWorkingShift(shift)) {
                        availableStaffs.push({
                            name: name,
                            shift: String(shift).trim(),
                            priority: st.priority
                        });
                    }
                });
 
                let workingNames = availableStaffs.map(s => s.name);
 
                // 担当者の決定：未設定を出さないため条件を段階的に緩めて必ず誰かを割当て、緩めた内容はエラーとして記載
                const KEY_NAMES = ["鍵No.1", "鍵No.2", "鍵No.3"];
                const plan = keyPlan.days[index];
                const dowK = dayData.dayOfWeek || '木';
                const holdPrev = [prevHolders["鍵No.1"], prevHolders["鍵No.2"], prevHolders["鍵No.3"]];
                const holdNow = [plan.slots[0], plan.slots[1], plan.slots[2] || fixedName];
                const storePrev = holdPrev.includes(STORE_NAME);
                const errs = [];
                const anyone = staffList.map(s => ({ name: s.name, shift: String((dayData.staffs && dayData.staffs[s.name]) || "公休").trim(), priority: s.priority }));
                const cntOf = (n, k) => { const s = staffList.find(x => x.name === n); return s ? s[k] : 0; };
                const rank = (list, task, ck) => list.slice().sort((a, b) => candScore(a.name, task, dateKey, dowK, cntOf(a.name, ck)) - candScore(b.name, task, dateKey, dowK, cntOf(b.name, ck)));
                const choose = (task, ck, pools, preds) => {
                    for (const p of pools) for (const f of [...preds, () => true]) for (const g of [true, false]) {
                        const l = rank(p.filter(x => f(x) && (!g || (allowed(x.name, task) && !ngHit(x.name, task, dateKey, dowK)))), task, ck);
                        if (l.length) return l[0];
                    }
                    return null;
                };
                const chkRule = (n, task, label) => {
                    if (!allowed(n, task)) errs.push(`${label}: ${n} は「担当可」がOFFの人です`);
                    if (ngHit(n, task, dateKey, dowK)) errs.push(`${label}: ${n} は個別ルール(NG)に該当します`);
                };

                // 1. 鍵開け：超早・早・フル。前日の夜までに鍵を持っている人
                const heldPrev = s => holdPrev.includes(s.name) || (RULES.storeOpen && storePrev);
                const openPool = availableStaffs.filter(s => canOpenShift(s.shift));
                const openCandidate = choose('open', 'openCount', [openPool, availableStaffs, anyone], RULES.openNeedsKey ? [heldPrev] : []);
                const openNameRaw = openCandidate ? openCandidate.name : UNSET;
                let openNoKey = false, openUsesStoreKey = false;
                if (openCandidate) {
                    if (!canOpenShift(openCandidate.shift)) errs.push(`鍵開け: 早番・超早・フルの出勤者がいないため ${openNameRaw}(${openCandidate.shift}) で代替（要確認）`);
                    openNoKey = RULES.openNeedsKey && !heldPrev(openCandidate);
                    if (openNoKey) errs.push(`鍵開け: ${openNameRaw} が前日の夜までに鍵を持っていません`);
                    openUsesStoreKey = RULES.storeOpen && storePrev && !holdPrev.includes(openNameRaw);
                    chkRule(openNameRaw, 'open', '鍵開け');
                }

                // 2. 鍵閉め：遅番・フル。退勤時に鍵を持っている人
                const accessC = s => holdNow.includes(s.name) || holdPrev.includes(s.name);
                const latePool = availableStaffs.filter(s => isLateShift(s.shift));
                const closeCandidate = choose('close', 'closeCount', [latePool, availableStaffs, anyone], [accessC]);
                const closeNameRaw = closeCandidate ? closeCandidate.name : UNSET;
                if (closeCandidate) {
                    if (!isLateShift(closeCandidate.shift)) errs.push(`鍵閉め: 遅番・フルの出勤者がいないため ${closeNameRaw}(${closeCandidate.shift}) で代替（要確認）`);
                    if (!accessC(closeCandidate)) errs.push(`鍵閉め: ${closeNameRaw} は退勤時に鍵を持っていません`);
                    chkRule(closeNameRaw, 'close', '鍵閉め');
                }

                // 3. 金庫チェック：遅番・フル。直近の担当者は避ける
                const recent = safeHist.slice(-Math.max(RULES.safeGap, 1));
                const noConsec = s => !RULES.safeNoConsec || !recent.includes(s.name);
                const notClose = s => s.name !== closeNameRaw;
                const safePool = availableStaffs.filter(s => canSafeShift(s.shift));
                const safeCandidate = choose('safe', 'safeCount', RULES.safeLateOnly ? [safePool] : [safePool, availableStaffs, anyone], [s => noConsec(s) && notClose(s), noConsec, notClose]);
                const safeNameRaw = safeCandidate ? safeCandidate.name : UNSET;
                const safeNote = `【金庫チェック連続】直近${recent.length}日以内に担当済み(${safeNameRaw})`;
                const isSafeContinuous = !!safeCandidate && RULES.safeNoConsec && recent.includes(safeNameRaw);
                if (safeCandidate) {
                    if (isSafeContinuous) errs.push(safeNote);
                    if (!isLateShift(safeCandidate.shift)) errs.push(`金庫チェック: 遅番・フルの出勤者がいないため ${safeNameRaw}(${safeCandidate.shift}) で代替（要確認）`);
                    chkRule(safeNameRaw, 'safe', '金庫チェック');
                    safeHist.push(safeNameRaw);
                }
 
                // ---------------------------------------------------------------
                // 鍵保持者（当日の退勤時点の保持者）… planKeyHolders で月全体を見て決定済み
                //  ・店舗置きは「翌日の鍵開け担当に他の方法で鍵が渡らない時」だけ
                //  ・休・応援・グランド等で出勤しない人は、鍵を持ったまま持ち帰り（店舗に置かない）
                // ---------------------------------------------------------------
                const newHolders = { "鍵No.1": plan.slots[0], "鍵No.2": plan.slots[1], "鍵No.3": plan.slots[2] || fixedName };
                KEY_NAMES.forEach(k => { if (newHolders[k] !== STORE_NAME) keyDays[newHolders[k]]++; });
 
                const k1Holder = newHolders["鍵No.1"];
                const k2Holder = newHolders["鍵No.2"];
                const k3Holder = newHolders["鍵No.3"];
 
                let keyHoldersResult = {
                    "鍵No.1": k1Holder,
                    "鍵No.2": k2Holder,
                    "鍵No.3": k3Holder
                };
 
                const openShift = openCandidate ? openCandidate.shift : "-";
                const openDisplay = openNameRaw !== UNSET ? `${openNameRaw}(${openShift})` : UNSET;
 
                const closeShift = closeCandidate ? closeCandidate.shift : "-";
                const closeDisplay = closeNameRaw !== UNSET ? `${closeNameRaw}(${closeShift})` : UNSET;
 
                const safeShift = safeCandidate ? safeCandidate.shift : "-";
                const safeDisplay = safeNameRaw !== UNSET ? `${safeNameRaw}(${safeShift})` : UNSET;
 
                let reasonsList = errs.slice();
 
                let storeKeys = [];
                if (k1Holder === STORE_NAME) storeKeys.push("鍵No.1");
                if (k2Holder === STORE_NAME) storeKeys.push("鍵No.2");
                if (k3Holder === STORE_NAME) storeKeys.push("鍵No.3");
 
                if (openCandidate && openNameRaw !== UNSET) {
                    const s = staffList.find(st => st.name === openNameRaw);
                    if (s) s.openCount++;
                }
                if (closeCandidate && closeNameRaw !== UNSET) {
                    const s = staffList.find(st => st.name === closeNameRaw);
                    if (s) s.closeCount++;
                }
                if (safeCandidate && safeNameRaw !== UNSET && safeNameRaw !== "-") {
                    const s = staffList.find(st => st.name === safeNameRaw);
                    if (s) {
                        s.safeCount++;
                        lastSafeStaff = safeNameRaw;
                    }
                }
 
                let checkStatus = errs.length === 0 ? "OK" : "NG";
                let okParts = [];
                if (openUsesStoreKey) okParts.push("鍵開けは前日に店舗へ置かれた鍵を使用");
                if (storeKeys.length > 0) okParts.push(`${storeKeys.join("・")}は店舗置き（翌日の鍵開け用）`);
                let okNote = okParts.length > 0 ? okParts.join(" / ") : "持ち帰り（問題なし）";
                let finalNote = checkStatus === "OK"
                    ? (reasonsList.length > 0 ? reasonsList.join(" / ") : okNote)
                    : `割当不可: ${reasonsList.join(" / ")}`;
 
                // シフト状況は「超早」「早」「遅」「遅＋」など原文のまま正確に表示
                let staffShiftsSummary = staffList.map(st => {
                    let sVal = (dayData.staffs && dayData.staffs[st.name]) ? String(dayData.staffs[st.name]).trim() : "公休";
                    return `${st.name}:${sVal}`;
                }).join(' / ');
 
                let shiftsByStaff = {};
                staffList.forEach(st => {
                    shiftsByStaff[st.name] = (dayData.staffs && dayData.staffs[st.name]) ? String(dayData.staffs[st.name]).trim() : "公休";
                });
 
                newShiftResult.push({
                    date: dateKey,
                    dayOfWeek: dayData.dayOfWeek || '木',
                    shiftsSummary: staffShiftsSummary,
                    shiftsByStaff: shiftsByStaff,
                    openStaff: openDisplay,
                    closeStaff: closeDisplay,
                    safeStaff: safeDisplay,
                    openRaw: openNameRaw,
                    closeRaw: closeNameRaw,
                    safeRaw: safeNameRaw,
                    keyHolders: { ...keyHoldersResult },
                    prevKeyHolders: { ...prevHolders },
                    checkStatus: checkStatus,
                    isContinuousNG: isSafeContinuous,
                    note: finalNote,
                    errs: errs,
                    locked: false
                });
 
                prevHolders = { ...keyHoldersResult };
            });
 
            shiftResult = newShiftResult;
        }
 
        // ---- 見やすさ改善用の表示ヘルパー ----
         
        // 人ごとの色（全画面で共通）。soft=薄い札（引渡者・見出し）／solid=濃い札（担当・引受者）
        const PERSON_STYLE = {
            "五家": { soft: "bg-slate-100 text-slate-700 border border-slate-300",  solid: "bg-indigo-100 text-indigo-900 border-2 border-indigo-500" },
            "爲貝": { soft: "bg-slate-100 text-slate-700 border border-slate-300", solid: "bg-indigo-100 text-indigo-900 border-2 border-indigo-500" },
            "柿沼": { soft: "bg-slate-100 text-slate-700 border border-slate-300",    solid: "bg-indigo-100 text-indigo-900 border-2 border-indigo-500" },
            "村瀬": { soft: "bg-slate-100 text-slate-700 border border-slate-300",    solid: "bg-indigo-100 text-indigo-900 border-2 border-indigo-500" },
            "楠美": { soft: "bg-slate-100 text-slate-700 border border-slate-300",  solid: "bg-indigo-100 text-indigo-900 border-2 border-indigo-500" }
        };
        const STORE_STYLE = { soft: "bg-white text-slate-500 border border-dashed border-slate-400", solid: "bg-white text-slate-700 border-2 border-dashed border-indigo-400" };
 
        function personPill(name, mode, size) {
            const st = (name === STORE_NAME) ? STORE_STYLE : (PERSON_STYLE[name] || { soft: "bg-slate-100 text-slate-700 border border-slate-300", solid: "bg-indigo-100 text-indigo-900 border-2 border-indigo-500" });
            const sz = size || "text-sm";
            return `<span class="inline-block min-w-[3.5rem] px-2.5 py-1 rounded-md ${sz} font-bold ${st[mode || 'soft']}">${name}</span>`;
        }
 
        function shiftBadge(val) {
            const v = String(val || "公休").trim();
            let cls;
            if (!isWorkingShift(v)) cls = "text-slate-400";
            else if (v.includes("超早")) cls = "bg-blue-200 text-blue-900";
            else if (v.includes("早") || isEarlyByTime(v)) cls = "bg-sky-100 text-sky-900";
            else if (isFullShift(v)) cls = "bg-violet-100 text-violet-900";
            else if (v.includes("遅")) cls = "bg-amber-100 text-amber-900";
            else cls = "bg-slate-100 text-slate-700";
            return `<span class="inline-block min-w-[3rem] px-2 py-1 rounded-md text-xs font-bold ${cls}">${v}</span>`;
        }
 
        // 担当者セル：名前（人ごとの色）＋その日のシフト
        function staffCellHtml(name, item) {
            if (!name || name === UNSET) return `<span class="text-sm font-bold">${UNSET}</span>`;
            const sv = (item.shiftsByStaff && item.shiftsByStaff[name]) ? item.shiftsByStaff[name] : "";
            return `<div>${personPill(name, 'plain')}</div><div class="mt-1">${sv ? shiftBadge(sv) : ""}</div>`;
        }
 
        const KEY_BORDER = { "鍵No.1": "border-red-400", "鍵No.2": "border-amber-400", "鍵No.3": "border-blue-400" };
 
        // 引渡者セル（前日から持っている人＝薄い札）
        function prevHolderCell(keyName, name) {
            return `<td class="px-2 py-3 text-center whitespace-nowrap border-l-2 border-slate-300">${personPill(name, 'soft')}</td>`;
        }
 
        // 引受者セル（今日受け取る人＝濃い札。同じ人なら「―」）
        function nextHolderCell(keyName, name) {
            if (!name) return `<td class="px-2 py-3 text-center whitespace-nowrap"><span class="text-slate-300 font-bold">―</span></td>`;
            return `<td class="px-2 py-3 text-center whitespace-nowrap">${personPill(name, 'solid')}</td>`;
        }
 
        function renderResult() {
            try {
                renderStaffHeaders();
                if (!shiftResult || shiftResult.length === 0) {
                    generateShift();
                }
                if (shiftResult.length > 0) {
                    document.getElementById('summaryPeriod').textContent = `${detectedExcelTitle} (${shiftResult[0].date} 〜 ${shiftResult[shiftResult.length - 1].date})`;
                }
 
                const ngCount = shiftResult.filter(r => r && getDisplayStatus(r).status === "NG").length;
                const warnEl = document.getElementById('summaryWarnings');
                if (warnEl) {
                    warnEl.textContent = ngCount + "件";
                    warnEl.className = ngCount > 0 ? "text-base font-bold text-amber-600" : "text-base font-bold text-emerald-600";
                }
 
                const statsBody = document.getElementById('statsTableBody');
                if (statsBody) {
                    statsBody.innerHTML = "";
                    staffList.sort((a, b) => a.priority - b.priority).forEach(s => {
                        const total = s.openCount + s.closeCount + s.safeCount;
                        const tr = document.createElement('tr');
                        tr.innerHTML = `
                            <td class="px-3 py-2 whitespace-nowrap text-center">${personPill(s.name, 'soft')}${roleLabel(s)}</td>
                            <td class="px-3 py-2 text-center whitespace-nowrap"><span class="px-2 py-0.5 bg-slate-100 rounded text-xs font-semibold">${s.priority}位</span></td>
                            <td class="px-3 py-2 text-center font-medium whitespace-nowrap">${s.openCount}回</td>
                            <td class="px-3 py-2 text-center font-medium whitespace-nowrap">${s.closeCount}回</td>
                            <td class="px-3 py-2 text-center font-medium whitespace-nowrap">${s.safeCount}回</td>
                            <td class="px-3 py-2 text-center font-bold text-indigo-600 whitespace-nowrap">${total}回</td>
                        `;
                        statsBody.appendChild(tr);
                    });
                }
 
                const resultBody = document.getElementById('resultTableBody');
                if (resultBody) {
                    resultBody.innerHTML = "";
                    let prevHoldersScreen = { "鍵No.1": "", "鍵No.2": "", "鍵No.3": "" };
 
                    shiftResult.forEach((item, idx) => {
                        if (!item) return;
                        const disp = getDisplayStatus(item);
                        const isNG = disp.status === "NG";
                        const isContinuousNG = item.isContinuousNG;
                        const kh = item.keyHolders || { "鍵No.1": "", "鍵No.2": "", "鍵No.3": "" };
 
                        let k1Prev = idx === 0 ? (item.prevKeyHolders || kh)["鍵No.1"] : prevHoldersScreen["鍵No.1"];
                        let k2Prev = idx === 0 ? (item.prevKeyHolders || kh)["鍵No.2"] : prevHoldersScreen["鍵No.2"];
                        let k3Prev = idx === 0 ? (item.prevKeyHolders || kh)["鍵No.3"] : prevHoldersScreen["鍵No.3"];
 
                        // 引渡者と引受者が同じ場合は引受者セルを空欄にする
                        let k1CurrentShow = (k1Prev === kh["鍵No.1"]) ? "" : kh["鍵No.1"];
                        let k2CurrentShow = (k2Prev === kh["鍵No.2"]) ? "" : kh["鍵No.2"];
                        let k3CurrentShow = (k3Prev === kh["鍵No.3"]) ? "" : kh["鍵No.3"];
 
                        const isOpenUnset = item.openRaw === UNSET;
                        const isCloseUnset = item.closeRaw === UNSET;
                        const isSafeUnset = item.safeRaw === UNSET;
 
                        const openCellClass = isOpenUnset ? "px-2 py-3 text-center text-red-700 font-bold border-2 border-red-400 whitespace-nowrap" : "px-2 py-3 text-center text-slate-800 whitespace-nowrap border-l-2 border-slate-300";
                        const closeCellClass = isCloseUnset ? "px-2 py-3 text-center text-red-700 font-bold border-2 border-red-400 whitespace-nowrap" : "px-2 py-3 text-center text-slate-800 whitespace-nowrap";
 
                        let safeCellClass = "px-2 py-3 text-center text-slate-800 whitespace-nowrap";
                        if (isSafeUnset) {
                            safeCellClass = "px-2 py-3 text-center text-red-700 font-bold border-2 border-red-400 whitespace-nowrap";
                        } else if (isContinuousNG) {
                            safeCellClass = "px-2 py-3 text-center text-orange-800 border-2 border-slate-600 whitespace-nowrap";
                        }
 
                        const dayColor = item.dayOfWeek === "日" ? "text-red-600" : (item.dayOfWeek === "土" ? "text-blue-600" : "text-slate-700");
                        let shiftCells = orderedStaff().map(s => s.name).map((n, i) => {
                            const v = (item.shiftsByStaff && item.shiftsByStaff[n]) ? item.shiftsByStaff[n] : "公休";
                            return `<td class="px-1.5 py-3 text-center whitespace-nowrap${i === 0 ? ' border-l-2 border-slate-300' : ''}">${shiftBadge(v)}</td>`;
                        }).join('');
 
                        const tr = document.createElement('tr');
                        tr.className = "hover:bg-slate-50 [&>td]:border-b [&>td]:border-slate-200";
                        tr.innerHTML = `
                            <td class="sticky left-0 z-10 bg-white px-3 py-3 text-center whitespace-nowrap border-r border-slate-200"><div class="text-sm font-bold ${dayColor}">${item.date}</div><div class="text-[11px] ${dayColor}">(${item.dayOfWeek})</div></td>
                            ${shiftCells}
                            <td class="${openCellClass}">${staffCellHtml(item.openRaw, item)}</td>
                            <td class="${closeCellClass}">${staffCellHtml(item.closeRaw, item)}</td>
                            <td class="${safeCellClass}">${staffCellHtml(item.safeRaw, item)}</td>
                            ${prevHolderCell("鍵No.1", k1Prev)}
                            ${nextHolderCell("鍵No.1", k1CurrentShow)}
                            ${prevHolderCell("鍵No.2", k2Prev)}
                            ${nextHolderCell("鍵No.2", k2CurrentShow)}
                            ${prevHolderCell("鍵No.3", k3Prev)}
                            ${nextHolderCell("鍵No.3", k3CurrentShow)}
                            <td class="px-2 py-3 text-center text-sm font-bold whitespace-nowrap border-l-2 border-slate-300 ${isNG ? 'text-amber-600' : 'text-emerald-600'}">${disp.status}</td>
                            <td class="px-2 py-3 text-center text-xs ${isNG ? 'text-amber-700 font-bold' : 'text-slate-500'}">${disp.note}</td>
                        `;
                        resultBody.appendChild(tr);
 
                        prevHoldersScreen = { ...kh };
                    });
                }
            } catch (err) {
                console.error("renderResult error:", err);
            }
        }
 
        // -------------------------------------------------------------------------
        // 画面4: 手動修正
        // -------------------------------------------------------------------------
        function renderEditTable() {
            try {
                const editBody = document.getElementById('editTableBody');
                if (!editBody) return;
                editBody.innerHTML = "";
                const staffNames = orderedStaff().map(s => s.name);

                // ヘッダー（日付｜シフト状況（人ごと）｜鍵開け…）を割り当て結果画面と同じ並びで作る
                const editHead = document.getElementById('editTableHead');
                if (editHead) {
                    const thBase = "px-3 py-2.5 font-semibold whitespace-nowrap text-center bg-slate-100";
                    const shiftNameThs = staffNames.map((n, i) =>
                        `<th class="px-2 py-1.5 text-[11px] whitespace-nowrap text-center bg-slate-100 text-slate-700${i === 0 ? ' border-l-2 border-slate-300' : ''}"><span class="px-2 py-0.5 rounded-md font-bold bg-white text-slate-800 border border-slate-300">${esc(n)}</span></th>`
                    ).join('');
                    editHead.innerHTML = `
                        <tr>
                            <th rowspan="2" class="${thBase} sticky left-0 z-20">日付</th>
                            <th colspan="${staffNames.length}" class="${thBase} border-l-2 border-slate-300">シフト状況</th>
                            <th rowspan="2" class="${thBase} border-l-2 border-slate-300">鍵開け</th>
                            <th rowspan="2" class="${thBase}">鍵閉め</th>
                            <th rowspan="2" class="${thBase}">金庫チェック</th>
                            <th colspan="2" class="px-3 py-2 font-semibold bg-red-50 text-slate-800 whitespace-nowrap text-center border-t-4 border-red-400 border-l-2 border-l-slate-300">鍵No.1 (赤)</th>
                            <th colspan="2" class="px-3 py-2 font-semibold bg-amber-50 text-slate-800 whitespace-nowrap text-center border-t-4 border-amber-400 border-l-2 border-l-slate-300">鍵No.2 (黄)</th>
                            <th colspan="2" class="px-3 py-2 font-semibold bg-blue-50 text-slate-800 whitespace-nowrap text-center border-t-4 border-blue-400 border-l-2 border-l-slate-300">鍵No.3 (青)</th>
                            <th rowspan="2" class="${thBase}">固定</th>
                            <th rowspan="2" class="${thBase} min-w-[180px]">判定（NG理由）</th>
                        </tr>
                        <tr>${shiftNameThs}${[['bg-red-50'], ['bg-amber-50'], ['bg-blue-50']].map(([c]) => `<th class="px-2 py-1.5 text-[11px] whitespace-nowrap text-center ${c} text-slate-600 border-l-2 border-slate-300">引渡者</th><th class="px-2 py-1.5 text-[11px] whitespace-nowrap text-center ${c} text-slate-600">引受者</th>`).join('')}</tr>`;
                }
 
                if (!shiftResult || shiftResult.length === 0) {
                    generateShift();
                }
 
                const dateMap = {};
                rawShiftData.forEach(item => {
                    if (!item || !item.date) return;
                    if (!dateMap[item.date]) dateMap[item.date] = {};
                    if (item.name) dateMap[item.date][item.name] = item.shiftType;
                });
 
                shiftResult.forEach((item, index) => {
                    if (!item) return;
                    const tr = document.createElement('tr');
                    const dayShifts = dateMap[item.date] || {};
                    const kh = item.keyHolders || { "鍵No.1": "", "鍵No.2": "", "鍵No.3": "" };
 
                    let makeOptions = (currentVal) => {
                        let opts = staffNames.map(n => {
                            let shiftName = dayShifts[n] || "公休";
                            let label = `${n} ${shiftName}`;
                            return `<option value="${n}" ${currentVal === n ? 'selected' : ''}>${label}</option>`;
                        }).join('');
                        opts += `<option value="${UNSET}" ${currentVal === UNSET ? 'selected' : ''}>${UNSET}</option>`;
                        return opts;
                    };
 
                    let makeSafeOptions = (currentVal) => {
                        let opts = staffNames.map(n => {
                            const sv = dayShifts[n] || "公休";
                            const ok = !RULES.safeLateOnly || canSafeShift(sv);
                            return `<option value="${n}" ${currentVal === n ? 'selected' : ''} ${ok || currentVal === n ? '' : 'disabled'}>${n} ${sv}${ok ? '' : ' ×'}</option>`;
                        }).join('');
                        opts += `<option value="${UNSET}" ${currentVal === UNSET ? 'selected' : ''}>${UNSET}</option>`;
                        return opts;
                    };

                    let makeHolderOptions = (currentVal) => {
                        let opts = staffNames.map(n => `<option value="${n}" ${currentVal === n ? 'selected' : ''}>${n}</option>`).join('');
                        opts += `<option value="${STORE_NAME}" ${currentVal === STORE_NAME ? 'selected' : ''}>${STORE_NAME}（店舗置き）</option>`;
                        return opts;
                    };
 
                    const editDayColor = item.dayOfWeek === "日" ? "text-red-600" : (item.dayOfWeek === "土" ? "text-blue-600" : "text-slate-700");
                    const editShiftCells = staffNames.map((n, i) => {
                        const v = (item.shiftsByStaff && item.shiftsByStaff[n]) ? item.shiftsByStaff[n] : (dayShifts[n] || "公休");
                        return `<td class="px-1.5 py-2.5 text-center whitespace-nowrap${i === 0 ? ' border-l-2 border-slate-300' : ''}">${shiftBadge(v)}</td>`;
                    }).join('');

                    tr.innerHTML = `
                        <td class="sticky left-0 z-10 bg-white px-3 py-2.5 font-medium whitespace-nowrap text-center border-r border-slate-200 ${editDayColor}">${item.date} (${item.dayOfWeek})</td>
                        ${editShiftCells}
                        <td class="px-3 py-2.5 text-center border-l-2 border-slate-300">
                            <select onchange="updateShiftField(${index}, 'openStaff', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeOptions(item.openRaw)}</select>
                        </td>
                        <td class="px-3 py-2.5 text-center">
                            <select onchange="updateShiftField(${index}, 'closeStaff', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeOptions(item.closeRaw)}</select>
                        </td>
                        <td class="px-3 py-2.5 text-center">
                            <select onchange="updateShiftField(${index}, 'safeStaff', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeSafeOptions(item.safeRaw)}</select>
                        </td>
                        ${editKeyCells(index)}
                        <td class="px-3 py-2.5 text-center">
                            <input type="checkbox" ${item.locked ? 'checked' : ''} onchange="toggleLock(${index}, this.checked)" class="w-4 h-4 text-indigo-600 rounded border-slate-300 cursor-pointer">
                        </td>
                        ${(() => { const d = getDisplayStatus(item); return `<td class="px-2 py-2.5 text-left text-[11px] ${d.status === 'NG' ? 'text-red-700 font-bold' : 'text-emerald-700'}"><span class="font-bold">${d.status}</span> ${d.status === 'NG' ? esc(String(d.note).replace(/^割当不可: /, '')) : ''}</td>`; })()}
                    `;
                    const ngRow = getDisplayStatus(item).status === 'NG';
                    if (ngRow) tr.classList.add('bg-red-50');
                    const ngOnlyEl = document.getElementById('editNgOnly');
                    if (ngOnlyEl && ngOnlyEl.checked && !ngRow) tr.classList.add('hidden');
                    editBody.appendChild(tr);
                });
            } catch (err) {
                console.error("renderEditTable error:", err);
            }
        }
 
        function updateShiftField(index, field, value) {
            if (!shiftResult[index]) return;
            if (field === 'openStaff') {
                shiftResult[index].openRaw = value;
                shiftResult[index].openStaff = value;
            }
            if (field === 'closeStaff') {
                shiftResult[index].closeRaw = value;
                shiftResult[index].closeStaff = value;
            }
            if (field === 'safeStaff') {
                shiftResult[index].safeRaw = value;
                shiftResult[index].safeStaff = value;
            }
            shiftResult[index].locked = true;
        }
 
        function updateHolderField(index, keyName, value) {
            if (!shiftResult[index]) return;
            if (!shiftResult[index].keyHolders) {
                shiftResult[index].keyHolders = { "鍵No.1": "", "鍵No.2": "", "鍵No.3": "" };
            }
            shiftResult[index].keyHolders[keyName] = value;
            shiftResult[index].locked = true;
        }
 
        function toggleLock(index, isLocked) {
            if (shiftResult[index]) {
                shiftResult[index].locked = isLocked;
            }
        }
 
        // -------------------------------------------------------------------------
        // 標準形式 Excel出力
        // -------------------------------------------------------------------------
        function exportToExcel() {
            if (!shiftResult || shiftResult.length === 0) {
                alert("出力するシフトデータがありません。先にシフトを作成してください。");
                return;
            }
 
            let html = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
            <head>
            <meta charset="UTF-8">
            <style>
                table { border-collapse: collapse; width: 100%; font-family: 'Meiryo', sans-serif; font-size: 11pt; }
                th, td { border: 1px solid #cbd5e1; padding: 6px 10px; text-align: center; vertical-align: middle; }
                .title { background-color: #4f46e5; color: #ffffff; font-size: 13pt; font-weight: bold; text-align: left; padding: 10px; }
                .header-col { background-color: #f1f5f9; color: #334155; font-weight: bold; }
                .cell-k1 { background-color: #fef2f2; color: #991b1b; }
                .cell-k2 { background-color: #fffbeb; color: #92400e; }
                .cell-k3 { background-color: #eff6ff; color: #1e40af; }
                .cell-unset { background-color: #fee2e2 !important; font-weight: bold; color: #991b1b; }
                .cell-continuous-ng { background-color: #fed7aa !important; font-weight: bold; color: #7c2d12; }
            </style>
            </head>
            <body>
            <table>
                <tr>
                    <td colspan="10" class="title">【${detectedExcelTitle} - 担当割り当て表】</td>
                </tr>
                <tr class="header-col">
                    <th>日付</th>
                    <th>シフト状況 (${orderedStaff().map(s => s.name).join('/')})</th>
                    <th>鍵開け</th>
                    <th>鍵閉め</th>
                    <th>金庫チェック担当</th>
                    <th>鍵1保持者</th>
                    <th>鍵2保持者</th>
                    <th>鍵3保持者</th>
                    <th>最終チェック</th>
                    <th>詳細理由</th>
                </tr>
            `;
 
            shiftResult.forEach((r, idx) => {
                if (!r) return;
                const kh = r.keyHolders || { "鍵No.1": "", "鍵No.2": "", "鍵No.3": "" };
                const isContinuousNG = r.isContinuousNG;
                const disp = getDisplayStatus(r);
 
                const isOpenUnset = r.openRaw === UNSET;
                const isCloseUnset = r.closeRaw === UNSET;
                const isSafeUnset = r.safeRaw === UNSET;
 
                const openStyle = isOpenUnset ? 'class="cell-unset"' : '';
                const closeStyle = isCloseUnset ? 'class="cell-unset"' : '';
 
                let safeStyle = '';
                if (isSafeUnset) {
                    safeStyle = 'class="cell-unset"';
                } else if (isContinuousNG) {
                    safeStyle = 'class="cell-continuous-ng"';
                }
 
                html += `
                <tr>
                    <td>${r.date} (${r.dayOfWeek})</td>
                    <td style="font-size: 10pt; background-color: #f8fafc; color: #334155;">${r.shiftsSummary || '-'}</td>
                    <td ${openStyle}>${r.openStaff}</td>
                    <td ${closeStyle}>${r.closeStaff}</td>
                    <td ${safeStyle}>${r.safeStaff}</td>
                    <td class="cell-k1">${kh["鍵No.1"]}</td>
                    <td class="cell-k2">${kh["鍵No.2"]}</td>
                    <td class="cell-k3">${kh["鍵No.3"]}</td>
                    <td style="font-weight: bold; color: ${disp.status === 'NG' ? '#d97706' : '#059669'};">${disp.status}</td>
                    <td style="text-align: left; color: #475569;">${disp.note}</td>
                </tr>`;
            });
 
            html += `</table></body></html>`;
 
            const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = "鍵当番・金庫チェック担当割当表_システム連動.xls";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
 
        // -------------------------------------------------------------------------
        // 添付Excel完全準拠形式出力（10月鍵開け2.xlsxの5,6,7行目・フォントサイズ・引渡者/引受者構造・同値時空欄）
        // -------------------------------------------------------------------------
        function exportToAttachedExcelStyle() {
            if (!shiftResult || shiftResult.length === 0) {
                alert("出力するシフトデータがありません。先にシフトを作成してください。");
                return;
            }
 
            let html = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
            <head>
            <meta charset="UTF-8">
            <style>
                table { border-collapse: collapse; width: 100%; font-family: 'HG創英角ｺﾞｼｯｸUB', 'Meiryo', sans-serif; }
                th, td { border: 1px solid #000000; padding: 4px 6px; text-align: center; vertical-align: middle; }
                .col-a { width: 160px; font-size: 20pt; font-family: 'MS PGothic', sans-serif; }
                .red-title { background-color: #ff0000; color: #ffffff; font-weight: bold; font-size: 36pt; text-align: left; font-family: 'HG創英角ｺﾞｼｯｸUB', sans-serif; }
                .note-text { text-align: left; font-size: 20pt; font-family: 'HG創英角ｺﾞｼｯｸUB', sans-serif; }
                .header-main { background-color: #f2f2f2; font-weight: bold; font-size: 18pt; font-family: 'HG創英角ｺﾞｼｯｸUB', sans-serif; }
                .sub-header { background-color: #e6e6e6; font-weight: bold; font-size: 18pt; font-family: 'HG創英角ｺﾞｼｯｸUB', sans-serif; }
                .data-cell { font-size: 24pt; font-family: 'MS PGothic', sans-serif; }
            </style>
            </head>
            <body>
            <table>
                <!-- 行1: タイトル -->
                <tr>
                    <td colspan="13" class="red-title">${detectedExcelTitle}</td>
                </tr>
                <!-- 行2: 注記1 -->
                <tr>
                    <td colspan="13" class="note-text">＊防犯管理上、いつ、誰が、どの鍵を持っているのか明確にし、受渡し時には必ず両者（引渡者・引受者）が押印して下さい。</td>
                </tr>
                <!-- 行3: 注記2 -->
                <tr>
                    <td colspan="13" class="note-text">＊金庫の鍵は、必ず営業中は携行し、退勤時は持ち帰る運用として下さい。</td>
                </tr>
                <!-- 行4: 空行 -->
                <tr>
                    <td colspan="13" style="border: 1px solid #000;">&nbsp;</td>
                </tr>
                <!-- 行5: 鍵の種類ヘッダー -->
                <tr class="header-main">
                    <td rowspan="2" class="col-a">鍵の<br>種類<br>・No.</td>
                    <td colspan="4">鍵No.１<br>（キーホルダー赤）</td>
                    <td colspan="4">鍵No.2<br>（キーホルダー黄）</td>
                    <td colspan="4">鍵No.3<br>（キーホルダー青）</td>
                </tr>
                <!-- 行6: 引渡者・引受者ヘッダー -->
                <tr class="sub-header">
                    <td colspan="2">引渡者</td>
                    <td colspan="2">引受者</td>
                    <td colspan="2">引渡者</td>
                    <td colspan="2">引受者</td>
                    <td colspan="2">引渡者</td>
                    <td colspan="2">引受者</td>
                </tr>
            `;
 
            let prevHolders = { "鍵No.1": "", "鍵No.2": "", "鍵No.3": "" };
 
            shiftResult.forEach((r, idx) => {
                if (!r) return;
                const kh = r.keyHolders || { "鍵No.1": "", "鍵No.2": "", "鍵No.3": "" };
 
                let k1Current = kh["鍵No.1"] || "";
                let k2Current = kh["鍵No.2"] || "";
                let k3Current = kh["鍵No.3"] || "";
 
                let k1Prev = idx === 0 ? (r.prevKeyHolders || kh)["鍵No.1"] : prevHolders["鍵No.1"];
                let k2Prev = idx === 0 ? (r.prevKeyHolders || kh)["鍵No.2"] : prevHolders["鍵No.2"];
                let k3Prev = idx === 0 ? (r.prevKeyHolders || kh)["鍵No.3"] : prevHolders["鍵No.3"];
 
                // 「引渡者」と「引受者」が同じ場合は引受者セルを空欄にする
                let k1PrevDisplay = k1Prev;
                let k1CurrentDisplay = (k1Prev === k1Current) ? "" : k1Current;
 
                let k2PrevDisplay = k2Prev;
                let k2CurrentDisplay = (k2Prev === k2Current) ? "" : k2Current;
 
                let k3PrevDisplay = k3Prev;
                let k3CurrentDisplay = (k3Prev === k3Current) ? "" : k3Current;
 
                // 日付 "10/5" → "5日"
                let dayParts = String(r.date).split("/");
                let dayLabel = dayParts.length === 2 ? `${dayParts[1]}日` : r.date;
 
                html += `
                <tr>
                    <td class="col-a">${dayLabel}</td>
                    <td colspan="2" class="data-cell">${k1PrevDisplay}</td>
                    <td colspan="2" class="data-cell">${k1CurrentDisplay}</td>
                    <td colspan="2" class="data-cell">${k2PrevDisplay}</td>
                    <td colspan="2" class="data-cell">${k2CurrentDisplay}</td>
                    <td colspan="2" class="data-cell">${k3PrevDisplay}</td>
                    <td colspan="2" class="data-cell">${k3CurrentDisplay}</td>
                </tr>`;
 
                prevHolders["鍵No.1"] = k1Current;
                prevHolders["鍵No.2"] = k2Current;
                prevHolders["鍵No.3"] = k3Current;
            });
 
            html += `</table></body></html>`;
 
            const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${String((shiftResult[0] && String(shiftResult[0].date).split("/")[0]) || detectedMonth || "")}月鍵管理表予定表_添付形式.xls`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
 
        window.onload = function() {
            const now = new Date();
            const ySel = document.getElementById('selYear');
            const mSel = document.getElementById('selMonth');
            for (let y = now.getFullYear() - 3; y <= now.getFullYear() + 2; y++) {
                ySel.add(new Option(y, y));
            }
            mSel.add(new Option("--", ""));   // 月は未選択スタート
            for (let m = 1; m <= 12; m++) mSel.add(new Option(m, m));
            ySel.value = now.getFullYear();
            mSel.value = "";

            // 年月を変えたら、読込済みファイルを選択年月で再解析
            const reparse = () => { if (getSelectedYM().m) reparseLoaded(); };
            ySel.addEventListener('change', reparse);
            mSel.addEventListener('change', reparse);

            renderPriorityList();
            renderStaffHeaders();
        };
    </script>

<script>
// ===== 追加：ルール設定・個別ルール・エラー表示 =====
let RULES = {
    openNeedsKey: true, noLateToEarly: true, storeAllow: true, storeOpen: false,
    safeNoConsec: true, safeGap: 1, prioWeight: 1, balance: 2, earlyHour: 9,
    deferTop: true, deferTopN: 2,
    safeLookahead: false, penOpen: 1000, penClose: 300, beam: 150,
    openPreferOn: true, openPrefer: "五家", openPreferWeight: 60,
    safeLateOnly: true, safeFullOk: true,
    kw: { off: "公休,休日,有給,欠勤,方針,グランド,応援,店長会議,会議,ミーティング,MTG,研修", early: "早", late: "遅,L,Ｌ", full: "フル" },
    custom: []
};
const TASKS = [['open', '鍵開け'], ['close', '鍵閉め'], ['safe', '金庫チェック']];
const KINDS = [['ng', 'NG（担当しない）'], ['prefer', '優先する']];
const DOWS = [['', '毎日'], ...DOW.map(d => [d, d + '曜日'])];

const kwHas = (s, k) => RULES.kw[k].split(/[,、，\s]+/).filter(Boolean).some(w => s.includes(w));
function isWorkingShift(x) { const s = String(x).trim(); if (!s || ['nan', 'undefined', 'null'].includes(s)) return false; return !kwHas(s, 'off'); }
function isFullShift(x) { return kwHas(String(x).trim().replace(/ﾌﾙ/g, "フル"), 'full'); }
function isEarlyByTime(x) { const h = shiftStartHour(x); return h !== null && h <= RULES.earlyHour; }
function canOpenShift(x) { const s = String(x).trim(); return isWorkingShift(s) && (kwHas(s, 'early') || isFullShift(s) || isEarlyByTime(s)); }
function isLateShift(x) {
    const s = String(x).trim(); if (!isWorkingShift(s)) return false;
    if (isFullShift(s)) return true;
    if (kwHas(s, 'early') || isEarlyByTime(s)) return false;
    const h = shiftStartHour(s);
    return kwHas(s, 'late') || (h !== null && h > RULES.earlyHour);
}
function allowed(name, task) { const s = staffList.find(x => x.name === name); return !!s && !(s.can && s.can[task] === false); }
function ruleHit(c, name, task, date, dow) {
    if (c.who !== name || c.task !== task) return false;
    if (c.dow && c.dow !== dow) return false;
    const ds = toHalfWidth(c.date || "").split(/[,、，\s]+/).filter(Boolean);
    return ds.length === 0 || ds.includes(date);
}
function ngHit(name, task, date, dow) { return RULES.custom.some(c => c.kind === 'ng' && ruleHit(c, name, task, date, dow)); }
function candScore(name, task, date, dow, cnt) {
    const s = staffList.find(x => x.name === name);
    let sc = (s ? s.priority : 99) * RULES.prioWeight + (cnt || 0) * RULES.balance;
    RULES.custom.forEach(c => { if (ruleHit(c, name, task, date, dow)) sc += c.kind === 'ng' ? 10000 : -50; });
    if (task === 'safe' && dow === '月' && s && s.mondaySafe) sc -= 50;
    return sc;
}
function setRule(k, v, num) { RULES[k] = num ? Number(v) : v; shiftResult = []; if (k === 'deferTopN' || k === 'deferTop') { try { renderPriorityList(); } catch (e) {} } }
function setKw(k, v) { RULES.kw[k] = v; shiftResult = []; }
function setCan(i, k, v) { const s = staffList[i]; if (!s) return; s.can = s.can || {}; s.can[k] = v; shiftResult = []; }
function addCustom() { RULES.custom.push({ who: staffList[0] ? staffList[0].name : "", task: 'safe', kind: 'ng', dow: '', date: '' }); shiftResult = []; renderRulesPanel(); }
function delCustom(i) { RULES.custom.splice(i, 1); shiftResult = []; renderRulesPanel(); }
function setCustom(i, k, v) { if (RULES.custom[i]) RULES.custom[i][k] = v; shiftResult = []; }

function renderRulesPanel() {
    const el = document.getElementById('rulesPanel'); if (!el) return;
    const ic = "border border-slate-300 rounded px-2 py-1 text-xs bg-white";
    const cb = (k, l) => `<label class="flex items-start gap-1.5 cursor-pointer"><input type="checkbox" class="mt-0.5" ${RULES[k] ? 'checked' : ''} onchange="setRule('${k}',this.checked)"><span>${l}</span></label>`;
    const num = (k, l) => `<label class="flex items-center gap-1.5">${l}<input type="number" min="0" value="${RULES[k]}" onchange="setRule('${k}',this.value,1)" class="w-16 ${ic}"></label>`;
    const kw = (k, l) => `<label class="block">${l}<input type="text" value="${esc(RULES.kw[k])}" onchange="setKw('${k}',this.value)" class="w-full mt-0.5 ${ic}"></label>`;
    const so = (arr, cur, i, k) => `<select onchange="setCustom(${i},'${k}',this.value)" class="${ic}">${arr.map(([v, l]) => `<option value="${esc(v)}" ${v === cur ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>`;
    const who = staffList.map(s => [s.name, s.name]);
    const rows = RULES.custom.map((c, i) => `<div class="flex flex-wrap items-center gap-1.5 p-2 bg-slate-50 border border-slate-200 rounded">${so(who, c.who, i, 'who')}${so(TASKS, c.task, i, 'task')}${so(KINDS, c.kind, i, 'kind')}${so(DOWS, c.dow, i, 'dow')}<input type="text" value="${esc(c.date)}" placeholder="日付 例 10/5,10/6（空=全日）" onchange="setCustom(${i},'date',this.value)" class="w-44 ${ic}"><button type="button" onclick="delCustom(${i})" class="p-1.5 hover:bg-red-100 rounded text-red-500 cursor-pointer"><i class="fa-solid fa-trash"></i></button></div>`).join('');
    const box = "bg-white p-5 rounded-xl shadow-sm border border-slate-200 space-y-3 text-xs text-slate-600";
    const h = (t, ico) => `<h3 class="font-bold text-slate-700 text-sm flex items-center gap-2"><i class="fa-solid ${ico} text-indigo-600"></i>${t}</h3>`;
    el.innerHTML = `
      <div class="${box}">${h('受け渡し・鍵のルール', 'fa-sliders')}
        ${cb('openNeedsKey', '鍵開け担当は前日の夜までに鍵を持っていること')}
        ${cb('noLateToEarly', '遅番（21:30以降に閉め）から早番（19:00退勤）への受け渡しを不可にする')}
        ${cb('storeAllow', '鍵の店舗置きを許可する（OFFでも回らない場合は最終手段として使用）')}
        ${cb('storeOpen', '店舗に置いた鍵を翌朝の鍵開けに使ってよい')}
        <div class="flex flex-wrap items-center gap-4">${cb('safeNoConsec', '金庫チェックの連続担当をNGにする')}${num('safeGap', '避ける日数')}</div>
        <div class="flex flex-wrap items-center gap-4 p-2 rounded bg-indigo-50 border border-indigo-200">${cb('deferTop', 'NGを出すくらいなら、優先順位の上位者には鍵を持たせるのを後回しにし、その人の鍵はあらかじめ他の人へ渡しておく')}${num('deferTopN', '上位何名を後回しにするか')}</div>
        <div class="p-2 rounded bg-emerald-50 border border-emerald-200 space-y-2"><div class="font-bold text-emerald-800">自動修正で追加される条件（手動でもON/OFFできます）</div>
          ${cb('safeLookahead', '金庫チェックは翌日以降の遅番を先読みし、連日NGになりそうな人を前日に避ける')}
          <label class="flex items-start gap-1.5 cursor-pointer"><input type="checkbox" class="mt-0.5" ${(RULES.penOpen || 0) >= 300000 ? 'checked' : ''} onchange="setStrongKey(this.checked)"><span>鍵開け・鍵閉めで「鍵を持っていない」NGを最優先で避ける（受け渡し・店舗置きを増やしてでも鍵を回す）</span></label>
          ${num('beam', '鍵の持ち回しの探索幅（大きいほど丁寧・遅くなる）')}
        </div>
      </div>
      <div class="${box}">${h('割当バランス（数字が大きいほど重視）', 'fa-scale-balanced')}
        <div class="flex flex-wrap gap-4">${num('prioWeight', '優先順位の重み')}${num('balance', '担当回数の均等化の重み')}</div>
        <p class="text-[11px] text-slate-400">例：均等化を0にすると優先順位どおり、大きくすると全員に均等に回ります。各メンバーの「担当可」は左の一覧で設定できます。</p>
      </div>
      <div class="${box}">${h('シフト表記の判定（カンマ区切りで追加・変更）', 'fa-tags')}
        ${kw('off', '出勤しない表記')}${kw('early', '早番の表記（超早も含む）')}${kw('late', '遅番の表記')}${kw('full', 'フルの表記')}
        ${num('earlyHour', '開始時刻がこの時刻以前なら早番（時）')}
      </div>
      <div class="${box}">${h('個別ルール（人・担当・曜日・日付ごとのNG／優先）', 'fa-user-lock')}
        ${rows || '<p class="text-slate-400">個別ルールはありません。</p>'}
        <button type="button" onclick="addCustom()" class="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-medium px-3 py-1.5 rounded-md cursor-pointer"><i class="fa-solid fa-plus"></i> ルールを追加</button>
      </div>`;
}

const _rpl = renderPriorityList;
renderPriorityList = function () {
    _rpl();
    [...document.getElementById('priorityList').children].forEach((div, i) => {
        const s = staffList[i]; if (!s) return; s.can = s.can || {};
        div.insertAdjacentHTML('beforeend', `<div class="flex flex-wrap gap-3 text-xs text-slate-600 pt-1 border-t border-slate-200">担当可：${TASKS.map(([k, l]) => `<label class="flex items-center gap-1 cursor-pointer"><input type="checkbox" ${s.can[k] === false ? '' : 'checked'} onchange="setCan(${i},'${k}',this.checked)">${l}</label>`).join('')}</div>`);
    });
    renderRulesPanel();
};

function renderErrorPanel() {
    const el = document.getElementById('errorPanel'); if (!el) return;
    const bad = (shiftResult || []).filter(r => r && getDisplayStatus(r).status === 'NG');
    if (!bad.length) { el.innerHTML = `<div class="p-4 rounded-xl border border-emerald-300 bg-emerald-50 text-emerald-800 text-sm font-bold"><i class="fa-solid fa-circle-check"></i> エラーなし：全日程で担当者が設定され、全ルールを満たしています。</div>`; return; }
    const items = bad.map(r => {
        const msgs = [...(r.errs || []), keyDupMessage(r.keyHolders)].filter(Boolean);
        if (!msgs.length) msgs.push(getDisplayStatus(r).note);
        return `<li class="flex gap-2"><span class="font-bold whitespace-nowrap">${esc(r.date)}(${esc(r.dayOfWeek)})</span><span>${msgs.map(esc).join(' ／ ')}</span></li>`;
    }).join('');
    el.innerHTML = `<div class="p-4 rounded-xl border border-red-300 bg-red-50"><div class="font-bold text-red-700 mb-2 text-sm"><i class="fa-solid fa-triangle-exclamation"></i> エラー・要確認：${bad.length}日（ルール設定や手動修正で調整できます）</div><ul class="space-y-1 text-xs text-red-800 max-h-64 overflow-auto">${items}</ul></div>`;
}
const _rr = renderResult;
renderResult = function () {
    _rr(); renderErrorPanel();
    const ng = (shiftResult || []).filter(r => r && getDisplayStatus(r).status === 'NG').length;
    const st = document.getElementById('summaryStatus');
    if (st) { st.textContent = ng ? `エラーあり（${ng}日）` : '全日程OK'; st.className = 'text-base font-bold ' + (ng ? 'text-red-600' : 'text-emerald-600'); }
    document.querySelectorAll('#resultTableBody tr').forEach((tr, i) => {
        const r = shiftResult[i];
        if (r && getDisplayStatus(r).status === 'NG') { tr.classList.add('bg-red-50'); const f = tr.firstElementChild; if (f) f.classList.replace('bg-white', 'bg-red-100'); }
    });
};
renderRulesPanel();
</script>
<script>
// ===== STEP0：元のシフト表(Excel)を「編集用シフト一覧表」に変換し、STEP1に自動で入れる =====
// ・「【印刷】シフト」のようなシートから、月・日数・スタッフ・シフトを自動判定（11月以外でもOK）
// ・「○月シフト一覧表」形式のExcelを作成（ダウンロード可）し、そのままSTEP1に読み込む
let step0State = null;   // { fileName, sheetName, y, m, lastDay, aoa, merges, staffCount }

function s0Str(v) { return (v === undefined || v === null) ? "" : String(v).trim(); }
function s0Serial(y, m, d) { return Date.UTC(y, m - 1, d) / 86400000 + 25569; }

// 日付セル（Date か Excelシリアル値）→ 年月日。それ以外は null
const S0_NAME_HDR = ["担当者名", "氏名", "名前", "スタッフ名", "従業員名", "社員名"];
const S0_CD_HDR = ["CD", "ＣＤ", "コード", "ｺｰﾄﾞ", "社員番号", "社員CD", "従業員番号", "NO", "NO."];
function s0ToYMD(cell) {
    // 文字列の日付（「2026/12/1」「12/1」「12月1日」）にも対応（他の月のファイル用）
    if (typeof cell === "string") {
        const t = toHalfWidth(cell.trim()).replace(/\(.*?\)|（.*?）/g, "");
        let m = t.match(/^(\d{4})[\/年.\-](\d{1,2})[\/月.\-](\d{1,2})日?$/);
        if (m) return { y: +m[1], m: +m[2], d: +m[3] };
        m = t.match(/^(\d{1,2})[\/月](\d{1,2})日?$/);
        if (m && +m[1] >= 1 && +m[1] <= 12 && +m[2] >= 1 && +m[2] <= 31) {
            const now = new Date();
            let y = now.getFullYear();
            if (+m[1] < now.getMonth() + 1 - 6) y++; else if (+m[1] > now.getMonth() + 1 + 6) y--;
            return { y: y, m: +m[1], d: +m[2] };
        }
        return null;
    }
    if (cell instanceof Date) {
        const dt = new Date(cell.getTime() + 12 * 3600 * 1000);   // タイムゾーン誤差吸収（STEP1と同じ）
        return { y: dt.getFullYear(), m: dt.getMonth() + 1, d: dt.getDate() };
    }
    if (typeof cell === "number" && cell >= 40000 && cell <= 60000) {
        const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(cell) * 86400000);
        return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
    }
    return null;
}

// 1シート分の行データ（sheet_to_json の header:1 形式）を解析する。使えないシートは null
function s0Analyze(rows, sheetName) {
    if (!rows || rows.length === 0) return null;
    // 見出し行（「担当者名」「CD」「役職」）を探す
    let hdr = -1, cName = -1;
    for (let r = 0; r < Math.min(rows.length, 80) && hdr < 0; r++) {
        const row = rows[r]; if (!row) continue;
        for (let c = 0; c < row.length; c++) if (S0_NAME_HDR.includes(s0Str(row[c]).replace(/[\s　]/g, ""))) { hdr = r; cName = c; break; }
    }
    if (hdr < 0) return null;
    let cCD = -1, cRole = -1;
    for (let c = 0; c < cName; c++) {
        const v = s0Str(rows[hdr][c]);
        if (S0_CD_HDR.includes(v.replace(/[\s　]/g, "").toUpperCase())) cCD = c;
        if (v.replace(/[\s　]/g, "") === "役職") cRole = c;
    }
    if (cCD < 0) return null;

    // 日付行：見出し行の少し上〜すぐ下で、日付が最も多く並ぶ行
    let dCells = [];
    for (let r = 0; r <= Math.min(rows.length - 1, hdr + 1); r++) {
        const row = rows[r]; if (!row) continue;
        const found = [];
        for (let c = cName + 1; c < row.length; c++) {
            const p = s0ToYMD(row[c]);
            if (p) found.push({ c: c, y: p.y, m: p.m, d: p.d });
        }
        if (found.length > dCells.length) dCells = found;
    }
    if (dCells.length < 28) return null;

    // 対象の年月 = 日付の中で最も多い年月（31列ぶんあっても、翌月に入った分は捨てる）
    const cnt = {};
    dCells.forEach(f => { const k = f.y + "-" + f.m; cnt[k] = (cnt[k] || 0) + 1; });
    const topKey = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
    const Y = parseInt(topKey.split("-")[0], 10), M = parseInt(topKey.split("-")[1], 10);
    const lastDay = new Date(Y, M, 0).getDate();
    const dayCol = {};
    dCells.forEach(f => { if (f.y === Y && f.m === M && dayCol[f.d] === undefined) dayCol[f.d] = f.c; });

    // スタッフ行：CDが数字で、担当者名が入っている行。CD列の「###」で部門が区切られる
    const groups = [[]];
    const seen = new Set();
    for (let r = hdr + 1; r < rows.length; r++) {
        const row = rows[r]; if (!row) continue;
        const cd = row[cCD], nm = s0Str(row[cName]);
        if (s0Str(cd) === "###") { if (groups[groups.length - 1].length) groups.push([]); continue; }
        const isNum = (typeof cd === "number") || /^\d{4,}$/.test(s0Str(cd));
        if (!isNum || !nm || seen.has(nm)) continue;
        seen.add(nm);
        const shifts = [];
        for (let d = 1; d <= lastDay; d++) {
            const v = dayCol[d] === undefined ? "" : row[dayCol[d]];
            const t = s0Str(v);
            shifts.push(t === "0" ? "" : t);   // 数式の空欄結果「0」は空欄として扱う
        }
        groups[groups.length - 1].push({ cd: (typeof cd === "number") ? cd : parseInt(s0Str(cd), 10), role: cRole >= 0 ? s0Str(row[cRole]) : "", name: nm, shifts: shifts });
    }
    while (groups.length && groups[groups.length - 1].length === 0) groups.pop();
    const staffCount = groups.reduce((t, g) => t + g.length, 0);
    if (staffCount === 0) return null;
    return { sheetName: sheetName, y: Y, m: M, lastDay: lastDay, groups: groups, staffCount: staffCount };
}

// 解析結果 → 「○月シフト一覧表」の表（行列データ）
function s0BuildAoa(a) {
    const DOWS = ['日', '月', '火', '水', '木', '金', '土'];
    const W = 3 + a.lastDay;
    // 部門が5つ（店長／AV／家電／情報／通信）のときは、店長とAVをひとまとめにする
    let groups = a.groups.map(g => g.slice());
    let labels;
    if (groups.length === 5) {
        groups = [groups[0].concat(groups[1]), groups[2], groups[3], groups[4]];
        labels = ["【店長・AV部門】", "【家電部門】", "【情報部門】", "【通信部門】"];
    } else {
        labels = groups.map((g, i) => "【部門" + (i + 1) + "】");
    }
    const aoa = [], merges = [];
    aoa.push([a.m + "月シフト一覧表"]);
    const dateRow = ["", "", ""], dayRow = ["CD", "役職", "担当者名"], dowRow = ["", "", ""];
    for (let d = 1; d <= a.lastDay; d++) {
        dateRow.push(s0Serial(a.y, a.m, d));
        dayRow.push(d + "日");
        dowRow.push(DOWS[new Date(a.y, a.m - 1, d).getDay()]);
    }
    aoa.push(dateRow); aoa.push(dayRow); aoa.push(dowRow);
    for (let c = 0; c < 3; c++) merges.push({ s: { r: 2, c: c }, e: { r: 3, c: c } });
    groups.forEach((g, gi) => {
        merges.push({ s: { r: aoa.length, c: 0 }, e: { r: aoa.length, c: W - 1 } });
        aoa.push([labels[gi]]);
        g.forEach(p => aoa.push([p.cd, p.role, p.name].concat(p.shifts)));
    });
    return { aoa: aoa, merges: merges, width: W };
}

function s0MakeWorkbook(st) {
    const ws = XLSX.utils.aoa_to_sheet(st.aoa);
    for (let c = 3; c < st.aoa[1].length; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: 1, c: c })];
        if (cell) { cell.t = 'n'; cell.z = 'm/d'; }
    }
    ws['!merges'] = st.merges;
    ws['!cols'] = [{ wch: 9 }, { wch: 11 }, { wch: 14 }].concat(new Array(st.lastDay).fill({ wch: 8 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "シート1");
    return wb;
}

function s0Download() {
    if (!step0State) { alert("先に元のシフト表Excelを選んでください。"); return; }
    XLSX.writeFile(s0MakeWorkbook(step0State), step0State.m + "月シフト_編集用.xlsx");
}

function s0SetYM(y, m) {
    const ys = document.getElementById('selYear');
    if (![...ys.options].some(o => parseInt(o.value, 10) === y)) ys.add(new Option(y, y));
    ys.value = String(y);
    document.getElementById('selMonth').value = String(m);
}

function s0RenderPreview() {
    const box = document.getElementById('s0Result');
    if (!box) return;
    if (!step0State) { box.classList.add('hidden'); return; }
    box.classList.remove('hidden');
    const st = step0State;
    document.getElementById('s0Info').textContent =
        `${st.fileName} ／ シート「${st.sheetName}」 ／ ${st.y}年${st.m}月（${st.lastDay}日分） ／ スタッフ${st.staffCount}名`;
    let html = "";
    st.aoa.forEach((row, ri) => {
        if (row.length === 1 && ri > 0) { html += `<tr><td colspan="${st.width}" class="px-2 py-1 text-left font-bold bg-indigo-50 text-indigo-700 sticky left-0">${esc(row[0])}</td></tr>`; return; }
        if (ri === 0) { html += `<tr><td colspan="${st.width}" class="px-2 py-1 text-left font-bold text-slate-700">${esc(row[0])}</td></tr>`; return; }
        if (ri === 1) return;
        const tag = ri <= 3 ? "th" : "td";
        html += "<tr>" + row.map((v, ci) => `<${tag} class="px-2 py-1 whitespace-nowrap border border-slate-100 ${ri <= 3 ? 'bg-slate-100 font-semibold' : ''}">${esc(v === 0 ? "0" : v)}</${tag}>`).join("") + "</tr>";
    });
    document.getElementById('s0Preview').innerHTML = html;
}

// 読み込んだワークブックから、変換できるシートを探して変換する
function s0Convert(workbook, fileName) {
    const names = workbook.SheetNames.slice();
    const rank = n => (n === "【印刷】シフト") ? 0 : (n.includes("印刷") && n.includes("シフト")) ? 1 : n.includes("入力用") ? 2 : n.includes("シフト") ? 3 : 4;
    names.sort((a, b) => rank(a) - rank(b));
    for (const name of names) {
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, blankrows: true });
        const a = s0Analyze(rows, name);
        if (!a) continue;
        const built = s0BuildAoa(a);
        return { fileName: fileName, sheetName: name, y: a.y, m: a.m, lastDay: a.lastDay, staffCount: a.staffCount, aoa: built.aoa, merges: built.merges, width: built.width };
    }
    return null;
}

// 変換した表を、STEP1（シフト取込）に読み込む
function s0FeedStep1(st) {
    s0SetYM(st.y, st.m);
    const ctx = { y: st.y, m: st.m };
    lastLoaded = { rows: st.aoa, name: "シート1" };
    const det = detectDateRow(st.aoa, ctx);
    parseExcelExactLayout(st.aoa, "シート1", ctx, det);
}

document.getElementById('rawShiftFile').addEventListener('change', function (ev) {
    const input = ev.target;
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
            const st = s0Convert(wb, file.name);
            if (!st) {
                alert("シフト表のシートが見つかりませんでした。「担当者名」の見出しと日付の行があるシート（例：【印刷】シフト）を含むExcelを選んでください。");
                return;
            }
            step0State = st;
            s0RenderPreview();
            s0FeedStep1(st);
            if (!rawShiftData || rawShiftData.length === 0) {
                // STEP1側でメンバーが見つからなかった場合（STEP1がメッセージを出して初期化済み）
                const msg = document.getElementById('s0Msg');
                if (msg) msg.textContent = "変換はできましたが、登録メンバーが見つからず STEP1 には入りませんでした。画面2でメンバー名・「Excel上の表記」を確認してください。";
            } else {
                const msg = document.getElementById('s0Msg');
                if (msg) msg.textContent = `STEP1に読み込みました（${st.y}年${st.m}月）。「1. シフト取込」へ進めます。編集用Excelはダウンロードもできます。`;
            }
        } catch (err) {
            console.error(err);
            alert("ファイルの解析に失敗しました。ファイル形式や内容を確認してください。");
        } finally {
            input.value = "";
        }
    };
    reader.readAsArrayBuffer(file);
});

// ===== リセット：読み込んだExcel・変換結果・割り当て結果を取り除く（メンバー・ルールの設定は残す） =====
function resetAll() {
    if (!confirm("読み込んだExcelファイルと、作成済みの割り当て結果をすべて取り除いてリセットします。\n（メンバー・優先順位・ルールの設定は残ります）\nよろしいですか？")) return;
    ['rawShiftFile', 'shiftFile'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ""; });
    step0State = null;
    const msg = document.getElementById('s0Msg'); if (msg) msg.textContent = "";
    s0RenderPreview();
    resetLoadedData();                       // 取込データ・割り当て結果・タイトルを初期化
    setLoadStatus("未読込（Excelを読み込んでください）", false);
    const ys = document.getElementById('selYear'), ms = document.getElementById('selMonth');
    if (ys) ys.value = String(new Date().getFullYear());
    if (ms) ms.value = "";
    try { renderResult(); } catch (e) { }
    try { renderEditTable(); } catch (e) { }
    const clear = (id, html) => { const el = document.getElementById(id); if (el) el.innerHTML = html; };
    clear('resultTableBody', ""); clear('editTableBody', ""); clear('errorPanel', "");
    const sp = document.getElementById('summaryPeriod'); if (sp) sp.textContent = "-";
    const ss = document.getElementById('summaryStatus'); if (ss) { ss.textContent = "-"; ss.className = "text-base font-bold text-slate-400"; }
    const sw = document.getElementById('summaryWarnings'); if (sw) { sw.textContent = "0件"; sw.className = "text-base font-bold text-emerald-600"; }
    switchTab('tab0');
}
</script>
<script>
// ===== 追加：NGの自動修正（3-1 → 3-2 → 3-3 …）・手動で一気に直す・結果の再検証 =====
let shiftVersions = [];      // { no, label, kind, changes[], rules, result, ngBefore }
let activeVer = -1;
let _fixing = false;         // 自動修正の試算中は版を作らない
let _pendingKind = null;     // 次の生成で作る版の種類（手動修正・再計算など）
let _shiftIdx = { dates: [], pos: {}, byDate: {}, dow: {} };
let fixMessage = "";

const deepClone = o => JSON.parse(JSON.stringify(o));

function buildShiftIndex() {
    const idx = { dates: [], pos: {}, byDate: {}, dow: {} };
    (rawShiftData || []).forEach(it => {
        if (!it || !it.date) return;
        if (!idx.byDate[it.date]) { idx.byDate[it.date] = {}; idx.pos[it.date] = idx.dates.length; idx.dates.push(it.date); idx.dow[it.date] = it.dayOfWeek; }
        if (it.name) idx.byDate[it.date][it.name] = it.shiftType;
    });
    _shiftIdx = idx;
}

// 金庫チェックの先読み：翌日以降に「その人しか遅番がいない」日があれば、今日の担当を避ける
const _candScoreBase = candScore;
candScore = function (name, task, date, dow, cnt) {
    let sc = _candScoreBase(name, task, date, dow, cnt);
    if (task === 'safe' && RULES.safeLookahead && RULES.safeNoConsec) {
        const i = _shiftIdx.pos[date];
        if (i !== undefined) {
            for (let j = 1; j <= Math.max(1, Number(RULES.safeGap) || 1); j++) {
                const d = _shiftIdx.dates[i + j]; if (!d) break;
                const sh = _shiftIdx.byDate[d] || {};
                const lates = staffList.filter(s => canSafeShift(sh[s.name] || "公休") && allowed(s.name, 'safe') && !ngHit(s.name, 'safe', d, _shiftIdx.dow[d]));
                if (lates.length === 1 && lates[0].name === name) { sc += 400; break; }
            }
        }
    }
    return sc;
};

// 金庫チェックができるシフトか（遅番。設定によりフルも可）
function canSafeShift(x) {
    const s = String(x || "").trim();
    if (!isLateShift(s)) return false;
    if (isFullShift(s) && RULES.safeFullOk === false) return false;
    return true;
}
const safeShiftLabel = () => (RULES.safeFullOk === false ? "遅番" : "遅番・フル");

// 全日程をルールに照らして再チェック（手動修正した日もOK/NGを正しく出す）
function validateAll(res) {
    let prev = null;
    const safeHist = [];
    const KEYS = ["鍵No.1", "鍵No.2", "鍵No.3"];
    (res || []).forEach((it, idx) => {
        if (!it) return;
        const day = _shiftIdx.byDate[it.date] || {};
        const sh = n => String((it.shiftsByStaff && it.shiftsByStaff[n]) || day[n] || "公休").trim();
        const kh = it.keyHolders || {};
        const pk = (idx === 0 || !prev) ? (it.prevKeyHolders || kh) : prev;
        it.prevKeyHolders = { ...pk };
        const holdPrev = KEYS.map(k => pk[k]), holdNow = KEYS.map(k => kh[k]);
        const storePrev = holdPrev.includes(STORE_NAME);
        const dow = it.dayOfWeek;
        const errs = [];
        const ruleChk = (n, task, label) => {
            if (!allowed(n, task)) errs.push(`${label}: ${n} は「担当可」がOFFの人です`);
            if (ngHit(n, task, it.date, dow)) errs.push(`${label}: ${n} は個別ルール(NG)に該当します`);
        };
        const disp = n => (n && n !== UNSET) ? `${n}(${sh(n)})` : UNSET;

        // 鍵開け
        if (!it.openRaw || it.openRaw === UNSET) errs.push("鍵開け: 担当者未設定");
        else {
            const s = sh(it.openRaw);
            if (!canOpenShift(s)) {
                const anyEarly = staffList.some(x => canOpenShift(sh(x.name)));
                errs.push(anyEarly ? `鍵開け: ${it.openRaw}(${s}) は早番・超早・フルではありません` : `鍵開け: 早番・超早・フルの出勤者がいないため ${it.openRaw}(${s}) で代替（要確認）`);
            }
            if (RULES.openNeedsKey && !holdPrev.includes(it.openRaw) && !(RULES.storeOpen && storePrev)) errs.push(`鍵開け: ${it.openRaw} が前日の夜までに鍵を持っていません`);
            ruleChk(it.openRaw, 'open', '鍵開け');
        }
        // 鍵閉め
        if (!it.closeRaw || it.closeRaw === UNSET) errs.push("鍵閉め: 担当者未設定");
        else {
            const s = sh(it.closeRaw);
            if (!isLateShift(s)) {
                const anyLate = staffList.some(x => isLateShift(sh(x.name)));
                errs.push(anyLate ? `鍵閉め: ${it.closeRaw}(${s}) は遅番・フルではありません` : `鍵閉め: 遅番・フルの出勤者がいないため ${it.closeRaw}(${s}) で代替（要確認）`);
            }
            if (!holdNow.includes(it.closeRaw) && !holdPrev.includes(it.closeRaw)) errs.push(`鍵閉め: ${it.closeRaw} は退勤時に鍵を持っていません`);
            ruleChk(it.closeRaw, 'close', '鍵閉め');
        }
        // 金庫チェック
        it.isContinuousNG = false;
        const anySafe = staffList.some(x => canSafeShift(sh(x.name)));
        if (!it.safeRaw || it.safeRaw === UNSET) errs.push(!anySafe && RULES.safeLateOnly ? `金庫チェック: ${safeShiftLabel()}の出勤者がいないため担当できる人がいません（${safeShiftLabel()}以外は金庫チェック不可）` : "金庫チェック: 担当者未設定");
        else if (it.safeRaw !== "-") {
            const s = sh(it.safeRaw);
            const recent = safeHist.slice(-Math.max(Number(RULES.safeGap) || 1, 1));
            if (RULES.safeNoConsec && recent.includes(it.safeRaw)) { it.isContinuousNG = true; errs.push(`【金庫チェック連続】直近${recent.length}日以内に担当済み(${it.safeRaw})`); }
            if (!canSafeShift(s)) {
                errs.push(RULES.safeLateOnly
                    ? `金庫チェック: ${it.safeRaw}(${s}) は${safeShiftLabel()}ではないため金庫チェックできません`
                    : (anySafe ? `金庫チェック: ${it.safeRaw}(${s}) は${safeShiftLabel()}ではありません` : `金庫チェック: ${safeShiftLabel()}の出勤者がいないため ${it.safeRaw}(${s}) で代替（要確認）`));
            }
            ruleChk(it.safeRaw, 'safe', '金庫チェック');
            safeHist.push(it.safeRaw);
        }
        // 鍵の受け渡し
        const present = x => x === STORE_NAME || isWorkingShift(sh(x));
        const lateOnly = x => x !== STORE_NAME && isLateShift(sh(x)) && !isFullShift(sh(x)) && !canOpenShift(sh(x));
        const earlyOnly = x => x !== STORE_NAME && canOpenShift(sh(x)) && !isFullShift(sh(x));
        KEYS.forEach(k => {
            const a = pk[k], b = kh[k];
            if (!a || !b || a === b) return;
            if (!present(a) || !present(b)) errs.push(`${k}: ${a}→${b} は出勤していない人がいるため受け渡しできません`);
            else if (RULES.noLateToEarly && lateOnly(a) && earlyOnly(b)) errs.push(`${k}: 遅番(${a})から早番(${b})へは受け渡しできません`);
        });

        it.openStaff = disp(it.openRaw); it.closeStaff = disp(it.closeRaw); it.safeStaff = disp(it.safeRaw);
        it.errs = errs;
        it.checkStatus = errs.length ? "NG" : "OK";
        if (errs.length) it.note = `割当不可: ${errs.join(" / ")}`;
        else {
            const ok = [];
            if (it.manualEdited) ok.push("手動で修正"); else if (it.autoAdjusted) ok.push("自動修正で調整"); else if (it.locked) ok.push("固定");
            if (RULES.storeOpen && storePrev && it.openRaw && !holdPrev.includes(it.openRaw)) ok.push("鍵開けは前日に店舗へ置かれた鍵を使用");
            const sk = KEYS.filter(k => kh[k] === STORE_NAME);
            if (sk.length) ok.push(`${sk.join("・")}は店舗置き（翌日の鍵開け用）`);
            it.note = ok.length ? ok.join(" / ") : "持ち帰り（問題なし）";
        }
        prev = { ...kh };
    });
}

function ngScore(res) {
    let days = 0, msgs = 0;
    (res || []).forEach(r => {
        if (!r) return;
        const d = getDisplayStatus(r);
        if (d.status === "NG") { days++; msgs += Math.max(1, (r.errs || []).length + (keyDupMessage(r.keyHolders) ? 1 : 0)); }
    });
    return { days, msgs, value: days * 1000 + msgs };
}

function recount(res) {
    staffList.forEach(s => { s.openCount = 0; s.closeCount = 0; s.safeCount = 0; });
    (res || []).forEach(r => {
        if (!r) return;
        const f = n => staffList.find(s => s.name === n);
        const o = f(r.openRaw), c = f(r.closeRaw), sf = f(r.safeRaw);
        if (o) o.openCount++; if (c) c.closeCount++; if (sf) sf.safeCount++;
    });
}

function pushVersion(kind, changes, ngBefore) {
    const no = shiftVersions.length + 1;
    shiftVersions.push({ no, label: `3-${no}`, kind, changes: changes || [], rules: deepClone(RULES), result: deepClone(shiftResult), ngBefore: ngBefore === undefined ? null : ngBefore });
    activeVer = shiftVersions.length - 1;
}

// 生成のたびに再チェックし、版（3-1, 3-2 …）を記録する
const _generateShiftBase = generateShift;
generateShift = function () {
    buildShiftIndex();
    _generateShiftBase();
    validateAll(shiftResult);
    // 金庫チェックの並びを月全体で見直し、連日NGなどが減る場合だけ採用
    if (shiftResult && shiftResult.length && RULES.safeNoConsec && typeof optimizeSafe === 'function') {
        const before = ngScore(shiftResult), saved = shiftResult.map(r => r.safeRaw);
        const path = optimizeSafe(shiftResult);
        if (path) {
            shiftResult.forEach((r, i) => { r.safeRaw = path[i]; });
            validateAll(shiftResult);
            if (ngScore(shiftResult).value < before.value) recount(shiftResult);
            else { shiftResult.forEach((r, i) => { r.safeRaw = saved[i]; }); validateAll(shiftResult); }
        }
    }
    if (_fixing) return;
    const kind = _pendingKind; _pendingKind = null;
    if (!shiftResult || shiftResult.length === 0) { shiftVersions = []; activeVer = -1; return; }
    if (!kind || shiftVersions.length === 0) { shiftVersions = []; fixMessage = ""; pushVersion("初回作成", []); }
    else pushVersion(kind.kind, kind.changes, kind.ngBefore);
};

recalculate = function () {
    try {
        if (!rawShiftData || rawShiftData.length === 0) { alert("先にExcelファイルを読み込んでください。"); return; }
        const before = ngScore(shiftResult).days;
        _pendingKind = shiftVersions.length ? { kind: "手動修正・再計算", changes: ["画面4での手動修正（固定した日）を反映して再計算"], ngBefore: before } : null;
        fixMessage = "";
        generateShift();
        renderResult();
        renderEditTable();
        const v = shiftVersions[activeVer];
        switchTab('tab3');   // 結果はポップアップで表示
    } catch (err) {
        console.error("Recalculate error:", err);
        alert("再計算中にエラーが発生しました。");
    }
};

// ---- 自動修正 ----
function evalWith(rules, base) {
    const saveR = RULES, saveS = shiftResult;
    RULES = rules; shiftResult = deepClone(base);
    _fixing = true;
    try { generateShift(); } finally { _fixing = false; }
    const res = shiftResult, sc = ngScore(res);
    RULES = saveR; shiftResult = saveS;
    return { res, sc };
}

function fixOptions(r) {
    const n = staffList.length, opts = [];
    if (!r.safeLookahead && r.safeNoConsec) opts.push(x => { x.safeLookahead = true; return "金庫チェックは翌日以降の遅番を先読みし、連日NGになりそうな人を前日に避ける"; });
    if ((r.penOpen || 1000) < 300000) opts.push(x => { x.penOpen = 300000; x.penClose = 300000; return "鍵開け・鍵閉めで「鍵を持っていない」NGを最優先で避ける（受け渡し・店舗置きを増やしてでも鍵を回す）"; });
    if (r.deferTop && (Number(r.beam) || 150) < 1200) opts.push(x => { x.beam = Math.min(1200, (Number(x.beam) || 150) * 3); return `鍵の持ち回しの探索幅を${x.beam}に広げ、より多くの組み合わせを調べる`; });
    if (r.deferTop) {
        const t = Number(r.deferTopN) || 0;
        if (t > 0) opts.push(x => { x.deferTopN = t - 1; return `鍵を持たせるのを後回しにする上位者を${t - 1}名に減らす`; });
        if (t < n - 3) opts.push(x => { x.deferTopN = t + 1; return `鍵を持たせるのを後回しにする上位者を${t + 1}名に増やす`; });
        opts.push(x => { x.deferTop = false; return "上位者の後回しをやめ、鍵No.3を固定担当で持ち回す方式にする"; });
    } else {
        opts.push(x => { x.deferTop = true; x.deferTopN = Math.max(0, Math.min(Number(x.deferTopN) || 2, n - 3)); return "上位者には鍵を持たせるのを後回しにし、休みの前に他の人へ渡しておく方式にする"; });
    }
    if (!r.storeAllow) opts.push(x => { x.storeAllow = true; return "鍵の店舗置きを許可する"; });
    if (!r.storeOpen) opts.push(x => { x.storeOpen = true; return "店舗に置いた鍵を翌朝の鍵開けに使ってよいとする"; });
    return opts;
}

function showBusy(on, text) {
    let el = document.getElementById('fixBusy');
    if (!el) {
        el = document.createElement('div');
        el.id = 'fixBusy';
        el.className = 'fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center hidden';
        el.innerHTML = '<div class="bg-white rounded-xl shadow-lg px-6 py-4 text-sm font-bold text-slate-700 flex items-center gap-3"><i class="fa-solid fa-spinner fa-spin text-indigo-600"></i><span id="fixBusyText"></span></div>';
        document.body.appendChild(el);
    }
    document.getElementById('fixBusyText').textContent = text || '';
    el.classList.toggle('hidden', !on);
}

// 金庫チェック担当を月全体で最適化（連日NG・遅番以外・個別NGを避け、変更は最小限）
function optimizeSafe(res) {
    const names = staffList.map(s => s.name);
    const g = Math.min(3, Math.max(1, Number(RULES.safeGap) || 1));
    let layer = new Map([["", { cost: 0, path: [] }]]);
    res.forEach(it => {
        const day = _shiftIdx.byDate[it.date] || {};
        const sh = n => String((it.shiftsByStaff && it.shiftsByStaff[n]) || day[n] || "公休").trim();
        let cands = (it.locked || !it.safeRaw || it.safeRaw === UNSET) ? [it.safeRaw] : names;
        if (!it.locked && RULES.safeLateOnly) { const ok = names.filter(n => canSafeShift(sh(n))); cands = ok.length ? ok : [it.safeRaw]; }
        const next = new Map();
        for (const [key, node] of layer) {
            const hist = key ? key.split("|") : [];
            cands.forEach(s => {
                let c = node.cost;
                if (s && s !== UNSET && s !== "-") {
                    if (!canSafeShift(sh(s))) c += 1000;
                    if (!allowed(s, 'safe') || ngHit(s, 'safe', it.date, it.dayOfWeek)) c += 1000;
                    if (RULES.safeNoConsec && hist.includes(s)) c += 1000;
                    if (s === it.closeRaw) c += 0.5;
                }
                if (s !== it.safeRaw) c += 1;
                const nh = (s && s !== UNSET && s !== "-") ? [...hist, s].slice(-g) : hist;
                const nk = nh.join("|");
                const ex = next.get(nk);
                if (!ex || c < ex.cost) next.set(nk, { cost: c, path: [...node.path, s] });
            });
        }
        layer = next;
    });
    let best = null;
    for (const v of layer.values()) if (!best || v.cost < best.cost) best = v;
    return best ? best.path : null;
}

// 1日ずつ担当・鍵保持者を組み替えて、全体のNGが減るものを採用（仕上げの微調整）
function repairDays(res) {
    const KEYS = ["鍵No.1", "鍵No.2", "鍵No.3"];
    const names = staffList.map(s => s.name);
    const holders = [...names, STORE_NAME];
    let score = (validateAll(res), ngScore(res));
    let adjusted = 0;
    // まず金庫チェックの並びを月全体で組み直す
    const path = optimizeSafe(res);
    if (path) {
        const saved = res.map(r => r.safeRaw);
        res.forEach((r, i) => { r.safeRaw = path[i]; });
        validateAll(res);
        const sc = ngScore(res);
        if (sc.value < score.value) { score = sc; res.forEach((r, i) => { if (r.safeRaw !== saved[i]) { r.autoAdjusted = true; adjusted++; } }); }
        else { res.forEach((r, i) => { r.safeRaw = saved[i]; }); validateAll(res); }
    }
    for (let pass = 0; pass < 3 && score.days > 0; pass++) {
        let improved = false;
        // NGの日と、その前日（金庫チェック連続などは前日を変えると直る）を対象にする
        const targets = new Set();
        res.forEach((r, i) => { if (r && getDisplayStatus(r).status === "NG") { targets.add(i); if (i > 0) targets.add(i - 1); } });
        res.forEach((it, i) => {
            if (!it || it.locked || !targets.has(i)) return;
            const save = { o: it.openRaw, c: it.closeRaw, s: it.safeRaw, k: { ...it.keyHolders } };
            const restore = () => { it.openRaw = save.o; it.closeRaw = save.c; it.safeRaw = save.s; it.keyHolders = { ...save.k }; };
            let best = null;
            const trial = () => {
                validateAll(res);
                const sc = ngScore(res);
                if (sc.value < score.value && (!best || sc.value < best.sc.value)) best = { sc, o: it.openRaw, c: it.closeRaw, s: it.safeRaw, k: { ...it.keyHolders } };
            };
            // 担当（鍵開け・鍵閉め・金庫チェック）の組み替え
            names.forEach(o => names.forEach(c => names.forEach(s => {
                it.openRaw = o; it.closeRaw = c; it.safeRaw = s; it.keyHolders = { ...save.k };
                trial();
            })));
            restore();
            // 鍵保持者を1本ずつ動かす（担当の組み替えと合わせて試す）
            const base = best ? { o: best.o, c: best.c, s: best.s } : { o: save.o, c: save.c, s: save.s };
            KEYS.forEach(k => holders.forEach(h => {
                it.openRaw = base.o; it.closeRaw = base.c; it.safeRaw = base.s;
                it.keyHolders = { ...(best ? best.k : save.k), [k]: h };
                trial();
            }));
            if (best) {
                it.openRaw = best.o; it.closeRaw = best.c; it.safeRaw = best.s; it.keyHolders = best.k;
                it.autoAdjusted = true; it.locked = true;
                score = best.sc; adjusted++; improved = true;
            } else restore();
            validateAll(res);
        });
        if (!improved) break;
    }
    validateAll(res);
    recount(res);
    return { res, sc: ngScore(res), adjusted };
}

// 修正ボタン：NGが無くなるか、これ以上減らせなくなるまで繰り返す
function autoFix() {
    if (!shiftResult || shiftResult.length === 0) { alert("先に割り当てを作成してください。"); return; }
    const start = ngScore(shiftResult);
    if (start.days === 0) { alert("NGはありません。"); return; }
    showBusy(true, `自動修正中…（NGが無くなるまで条件を追加して作り直しています）`);
    setTimeout(() => {
        const saveR = deepClone(RULES), saveS = deepClone(shiftResult);
        try {
            let rules = deepClone(RULES), res = deepClone(shiftResult), cur = start;
            const changes = [];
            for (let round = 0; round < 8 && cur.days > 0; round++) {
                RULES = rules;
                const opts = fixOptions(rules);
                let best = null;
                const tryRules = fns => {
                    const r = deepClone(rules);
                    const labels = fns.map(f => f(r));
                    const ev = evalWith(r, res);
                    if (!best || ev.sc.value < best.sc.value) best = { rules: r, labels, res: ev.res, sc: ev.sc };
                };
                opts.forEach(f => tryRules([f]));
                if (!best || best.sc.value >= cur.value) {
                    for (let i = 0; i < opts.length; i++) for (let j = i + 1; j < opts.length; j++) tryRules([opts[i], opts[j]]);
                }
                if (!best || best.sc.value >= cur.value) break;
                rules = best.rules; res = best.res; cur = best.sc;
                best.labels.forEach(l => changes.push(l));
            }
            // 仕上げ：残ったNGの日を1日ずつ組み替え
            RULES = rules;
            buildShiftIndex();
            const rep = repairDays(res);
            if (rep.sc.value < cur.value) { res = rep.res; cur = rep.sc; changes.push(`残ったNGの日（${rep.adjusted}日）の担当・鍵保持者を個別に組み替え`); }

            if (cur.value < start.value) {
                RULES = rules;
                shiftResult = res;
                recount(shiftResult);
                pushVersion(`自動修正${shiftVersions.filter(v => v.kind.startsWith("自動修正")).length + 1}回目`, changes, start.days);
                fixMessage = cur.days === 0
                    ? `✅ 「${shiftVersions[activeVer].label} 割り当て結果」でNGが無くなりました（${start.days}日 → 0日）。`
                    : `⚠ NGを ${start.days}日 → ${cur.days}日 まで減らしました。残りはシフト上どうしても担当できる人がいない日などです。「手動で一気に直す」から修正してください。`;
                try { renderPriorityList(); } catch (e) { }
            } else {
                RULES = saveR; shiftResult = saveS; recount(shiftResult);
                fixMessage = "⚠ 自動修正ではこれ以上NGを減らせませんでした（シフト上どうしても担当できる人がいない日など）。残りは「手動で一気に直す」から修正してください。";
            }
            renderResult();
        } catch (err) {
            console.error("autoFix error:", err);
            RULES = saveR; shiftResult = saveS; recount(shiftResult);
            alert("自動修正中にエラーが発生しました。");
        } finally {
            showBusy(false);
        }
    }, 30);
}

function goManualFix() {
    if (!shiftResult || shiftResult.length === 0) { alert("先に割り当てを作成してください。"); return; }
    switchTab('tab4');
    const cb = document.getElementById('editNgOnly');
    if (cb) { cb.checked = ngScore(shiftResult).days > 0; renderEditTable(); }
}

function selectVersion(i) {
    const v = shiftVersions[i]; if (!v) return;
    activeVer = i;
    RULES = deepClone(v.rules);
    shiftResult = deepClone(v.result);
    recount(shiftResult);
    fixMessage = "";
    try { renderPriorityList(); } catch (e) { }
    renderResult();
}

function renderFixPanel() {
    const el = document.getElementById('fixPanel'); if (!el) return;
    const lbl = document.getElementById('resultVerLabel');
    const v = shiftVersions[activeVer];
    if (lbl) lbl.textContent = v ? v.label : "3-1";
    if (!v || !shiftResult || shiftResult.length === 0) { el.innerHTML = ""; return; }
    const ng = ngScore(shiftResult).days;
    const chips = shiftVersions.map((x, i) => {
        const xn = ngScore(x.result).days;
        const on = i === activeVer;
        return `<button type="button" onclick="selectVersion(${i})" class="px-3 py-1.5 rounded-lg border text-xs font-bold cursor-pointer ${on ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'}">${x.label}<span class="font-normal ml-1">（${esc(x.kind)}）</span><span class="ml-1.5 px-1.5 rounded ${xn ? (on ? 'bg-red-500' : 'bg-red-100 text-red-700') : (on ? 'bg-emerald-500' : 'bg-emerald-100 text-emerald-700')}">NG ${xn}日</span></button>`;
    }).join('<i class="fa-solid fa-chevron-right text-slate-300 text-[10px]"></i>');
    const changes = v.changes.length
        ? `<div class="text-xs text-slate-600"><span class="font-bold">${v.label} で追加・変更した条件：</span><ul class="list-disc list-inside mt-1 space-y-0.5">${v.changes.map(c => `<li>${esc(c)}</li>`).join('')}</ul>${v.ngBefore !== null ? `<div class="mt-1">NG：${v.ngBefore}日 → <span class="font-bold">${ng}日</span></div>` : ''}</div>`
        : '';
    const next = `3-${shiftVersions.length + 1}`;
    const prevV = shiftVersions[activeVer - 1], nextV = shiftVersions[activeVer + 1];
    const navBtn = (cond, onclick, html) => `<button type="button" ${cond ? `onclick="${onclick}"` : 'disabled'} class="px-3 py-1.5 rounded-lg border text-xs font-bold ${cond ? 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50 cursor-pointer' : 'bg-slate-100 border-slate-200 text-slate-300 cursor-not-allowed'}">${html}</button>`;
    el.innerHTML = `
      <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-3">
        <div class="flex flex-wrap items-center gap-2">
          ${navBtn(!!prevV, `selectVersion(${activeVer - 1})`, `<i class="fa-solid fa-arrow-left"></i> 前の版に戻る${prevV ? `（${prevV.label}）` : ''}`)}
          ${navBtn(!!nextV, `selectVersion(${activeVer + 1})`, `次の版へ${nextV ? `（${nextV.label}）` : ''} <i class="fa-solid fa-arrow-right"></i>`)}
          ${navBtn(shiftVersions.length > 1, `openCompare()`, `<i class="fa-solid fa-table-columns"></i> 版を並べて比較する`)}
        </div>
        <div class="flex flex-wrap items-center gap-2"><span class="text-xs font-bold text-slate-500 mr-1"><i class="fa-solid fa-code-branch"></i> 版：</span>${chips}</div>
        ${changes}
        <div class="flex flex-wrap items-center gap-3">
          <button type="button" onclick="autoFix()" ${ng ? '' : 'disabled'} class="${ng ? 'bg-emerald-600 hover:bg-emerald-700 cursor-pointer' : 'bg-slate-300 cursor-not-allowed'} text-white font-semibold px-4 py-2 rounded-lg shadow text-sm inline-flex items-center gap-2"><i class="fa-solid fa-wand-magic-sparkles"></i> 自動で修正する（NGが無くなるまで → ${next} を作成）</button>
          <button type="button" onclick="goManualFix()" class="bg-white hover:bg-amber-50 border-2 border-amber-500 text-amber-700 font-semibold px-4 py-2 rounded-lg text-sm inline-flex items-center gap-2 cursor-pointer"><i class="fa-solid fa-hand"></i> ${ng ? '修正を飛ばして手動で一気に直す' : '手動で修正する'}</button>
          ${ng ? '' : `<span class="text-xs font-bold text-emerald-700"><i class="fa-solid fa-circle-check"></i> NGはありません</span>`}
        </div>
        ${fixMessage ? `<div class="text-xs font-bold ${fixMessage.startsWith('⚠') ? 'text-amber-700 bg-amber-50 border-amber-200' : 'text-emerald-700 bg-emerald-50 border-emerald-200'} border rounded-lg px-3 py-2">${esc(fixMessage)}</div>` : ''}
      </div>`;
}

// ===== 版の比較画面（左右に並べて、違うところを色付け） =====
let cmpLeft = -1, cmpRight = -1;

function openCompare() {
    if (shiftVersions.length < 2) { alert("比較するには版が2つ以上必要です。先に「自動で修正する」などで新しい版を作ってください。"); return; }
    cmpRight = activeVer;
    cmpLeft = activeVer > 0 ? activeVer - 1 : 1;
    switchTab('tabCmp');
}

function setCmp(side, v) { if (side === 'L') cmpLeft = Number(v); else cmpRight = Number(v); renderCompare(); }

function renderCompare() {
    const el = document.getElementById('cmpBody'); if (!el) return;
    if (shiftVersions.length < 2) {
        el.innerHTML = `<p class="text-slate-400 text-sm py-8 text-center">比較できる版がありません。「3. 割り当て結果」で「自動で修正する」を押すと、新しい版（3-2 …）が作られ、ここで並べて見比べられます。</p>`;
        return;
    }
    if (!shiftVersions[cmpLeft]) cmpLeft = 0;
    if (!shiftVersions[cmpRight]) cmpRight = shiftVersions.length - 1;
    const L = shiftVersions[cmpLeft], R = shiftVersions[cmpRight];
    const sel = (side, cur) => `<select onchange="setCmp('${side}', this.value)" class="border border-slate-300 rounded-md px-2 py-1 text-xs bg-white font-bold">${shiftVersions.map((x, i) => `<option value="${i}" ${i === cur ? 'selected' : ''}>${x.label}（${esc(x.kind)}）NG ${ngScore(x.result).days}日</option>`).join('')}</select>`;
    const head = (v, side, cur, color) => `
        <div class="flex-1 min-w-[16rem] p-3 rounded-lg border ${color} space-y-1.5">
          <div class="flex flex-wrap items-center gap-2"><span class="text-xs font-bold">${side === 'L' ? '左' : '右'}：</span>${sel(side, cur)}
            <button type="button" onclick="selectVersion(${cur}); switchTab('tab3');" class="text-xs px-2 py-1 rounded-md bg-indigo-600 hover:bg-indigo-700 text-white cursor-pointer">この版を表示・採用</button></div>
          <div class="text-xs">NG：<span class="font-bold ${ngScore(v.result).days ? 'text-red-600' : 'text-emerald-600'}">${ngScore(v.result).days}日</span></div>
          ${v.changes.length ? `<ul class="list-disc list-inside text-[11px] text-slate-500">${v.changes.map(c => `<li>${esc(c)}</li>`).join('')}</ul>` : ''}
        </div>`;
    const fields = [["openRaw", "鍵開け"], ["closeRaw", "鍵閉め"], ["safeRaw", "金庫"], ["k1", "鍵1"], ["k2", "鍵2"], ["k3", "鍵3"]];
    const val = (r, f) => !r ? "" : f === "k1" ? r.keyHolders["鍵No.1"] : f === "k2" ? r.keyHolders["鍵No.2"] : f === "k3" ? r.keyHolders["鍵No.3"] : r[f];
    const thL = fields.map((f, i) => `<th class="px-1.5 py-1.5 bg-slate-100 whitespace-nowrap${i === 0 ? ' border-l-2 border-slate-400' : ''}">${f[1]}</th>`).join('') + `<th class="px-1.5 py-1.5 bg-slate-100">判定</th>`;
    let diffDays = 0;
    const rows = L.result.map((lr, i) => {
        const rr = R.result[i];
        const ls = getDisplayStatus(lr), rs = rr ? getDisplayStatus(rr) : { status: "", note: "" };
        let rowDiff = false;
        const cells = (r, st, other, isRight) => fields.map((f, k) => {
            const v = val(r, f[0]), o = val(other, f[0]);
            const diff = v !== o;
            if (diff) rowDiff = true;
            return `<td class="px-1.5 py-1.5 whitespace-nowrap${k === 0 ? ' border-l-2 border-slate-400' : ''} ${diff ? (isRight ? 'bg-yellow-200 font-bold text-slate-900' : 'bg-yellow-50 text-slate-500 line-through decoration-slate-400') : ''}">${esc(v || '')}</td>`;
        }).join('') + `<td class="px-1.5 py-1.5 font-bold ${st.status === 'NG' ? 'text-red-600' : 'text-emerald-600'}" title="${esc(st.note)}">${st.status}${st.status === 'NG' ? ngDots(r) : ''}</td>`;
        const left = cells(lr, ls, rr, false), right = cells(rr, rs, lr, true);
        if (rowDiff) diffDays++;
        const dc = lr.dayOfWeek === "日" ? "text-red-600" : lr.dayOfWeek === "土" ? "text-blue-600" : "";
        const shifts = orderedStaff().map(s => `<td class="px-1 py-1.5">${shiftBadge((lr.shiftsByStaff || {})[s.name] || "公休")}</td>`).join('');
        return `<tr class="${rowDiff ? 'bg-yellow-50/40' : ''} border-b border-slate-200"><td class="sticky left-0 bg-white px-2 py-1.5 font-bold whitespace-nowrap ${dc}">${esc(lr.date)}(${esc(lr.dayOfWeek)})</td>${shifts}${left}${right}</tr>`;
    }).join('');
    el.innerHTML = `
      <div class="flex flex-wrap gap-3 mb-3">${head(L, 'L', cmpLeft, 'border-slate-300 bg-slate-50')}${head(R, 'R', cmpRight, 'border-indigo-300 bg-indigo-50')}</div>
      <div class="text-xs text-slate-600 mb-2">違いがある日：<span class="font-bold">${diffDays}日</span>　<span class="inline-block px-2 rounded bg-yellow-200 font-bold">黄色</span>＝右の版で変わった所（左の版の元の値は打ち消し線）。判定にマウスを乗せるとNG理由が出ます。</div>
      <div class="overflow-auto border border-slate-200 rounded-lg max-h-[70vh]">
        <table class="compact-table w-full text-xs text-center text-slate-700 border-separate border-spacing-0">
          <thead class="sticky top-0 z-10">
            <tr><th rowspan="2" class="sticky left-0 z-20 bg-slate-100 px-2">日付</th><th colspan="${orderedStaff().length}" class="bg-slate-100">シフト</th>
              <th colspan="${fields.length + 1}" class="bg-slate-200 border-l-2 border-slate-400">左：${L.label}</th>
              <th colspan="${fields.length + 1}" class="bg-indigo-100 border-l-2 border-slate-400">右：${R.label}</th></tr>
            <tr>${orderedStaff().map(s => `<th class="px-1 py-1.5 bg-slate-100 whitespace-nowrap">${esc(s.name)}</th>`).join('')}${thL}${thL}</tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
}

function openCompareTab() {
    if (shiftVersions.length >= 2 && (cmpLeft < 0 || cmpRight < 0)) { cmpRight = activeVer; cmpLeft = activeVer > 0 ? activeVer - 1 : 1; }
    switchTab('tabCmp');
}
const _switchTabBase = switchTab;
switchTab = function (tabId) { _switchTabBase(tabId); if (tabId === 'tabCmp') renderCompare(); };

const _renderResultWithFix = renderResult;
renderResult = function () { _renderResultWithFix(); renderFixPanel(); };

const _renderEditTableBase = renderEditTable;
renderEditTable = function () {
    _renderEditTableBase();
    const c = document.getElementById('editNgCount');
    if (c) { const n = ngScore(shiftResult).days; c.innerHTML = n ? `<span class="font-bold text-red-600">NG ${n}日</span>（赤い行）` : '<span class="font-bold text-emerald-600">NGなし</span>'; }
};

function setStrongKey(on) { RULES.penOpen = on ? 300000 : 1000; RULES.penClose = on ? 300000 : 300; shiftResult = []; }
</script>
<script>
// ===== 追加：NGの種類ごとの色分け =====
const NG_CATS = {
    key:   { label: "鍵の割り当て（受け渡し・保持）", chip: "bg-rose-100 text-rose-800 border-rose-300",       dot: "bg-rose-500",   cell: "!bg-rose-100" },
    shift: { label: "鍵開け・鍵閉めのシフト不足",     chip: "bg-violet-100 text-violet-800 border-violet-300", dot: "bg-violet-500", cell: "!bg-violet-100" },
    safe:  { label: "金庫チェック",                   chip: "bg-amber-100 text-amber-900 border-amber-400",    dot: "bg-amber-500",  cell: "!bg-amber-100" },
    rule:  { label: "担当可・個別ルール",             chip: "bg-sky-100 text-sky-800 border-sky-300",          dot: "bg-sky-500",    cell: "!bg-sky-100" },
    unset: { label: "担当者未設定",                   chip: "bg-red-600 text-white border-red-700",            dot: "bg-red-700",    cell: "!bg-red-200" }
};
const NG_ORDER = ["key", "shift", "safe", "rule", "unset"];

function ngCat(msg) {
    msg = String(msg || "");
    if (msg.includes("未設定")) return "unset";
    if (msg.includes("担当可") || msg.includes("個別ルール")) return "rule";
    if (msg.includes("金庫チェック")) return "safe";
    if (msg.startsWith("鍵No.") || msg.includes("鍵を持っていません") || msg.includes("重複")) return "key";
    return "shift";
}
function ngMsgs(r) {
    if (!r || getDisplayStatus(r).status !== "NG") return [];
    const m = [...(r.errs || []), keyDupMessage(r.keyHolders)].filter(Boolean);
    if (!m.length) m.push(String(getDisplayStatus(r).note).replace(/^割当不可: /, ""));
    return m;
}
function ngCatsOf(r) { return [...new Set(ngMsgs(r).map(ngCat))].sort((a, b) => NG_ORDER.indexOf(a) - NG_ORDER.indexOf(b)); }
function ngDots(r) { return `<span class="inline-flex gap-0.5 ml-1 align-middle">${ngCatsOf(r).map(c => `<span title="${esc(NG_CATS[c].label)}" class="inline-block w-2.5 h-2.5 rounded-full ${NG_CATS[c].dot}"></span>`).join('')}</span>`; }
function ngChip(msg) { const c = NG_CATS[ngCat(msg)]; return `<span class="inline-block px-1.5 py-0.5 rounded border text-[11px] font-bold ${c.chip}">${esc(msg)}</span>`; }
function ngLegend(res) {
    const cnt = {};
    (res || []).forEach(r => ngCatsOf(r).forEach(c => { cnt[c] = (cnt[c] || 0) + 1; }));
    return `<div class="flex flex-wrap items-center gap-2 text-[11px]"><span class="font-bold text-slate-500">NGの種類：</span>${NG_ORDER.map(c => `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full border ${NG_CATS[c].chip}"><span class="w-2.5 h-2.5 rounded-full ${NG_CATS[c].dot}"></span>${NG_CATS[c].label}${res ? `：${cnt[c] || 0}日` : ''}</span>`).join('')}</div>`;
}
// エラーがどの欄のものか（鍵開け・鍵閉め・金庫・鍵1〜3）
function ngTargets(msg) {
    msg = String(msg);
    if (msg.includes("重複")) return ["k1", "k2", "k3"];
    const k = msg.match(/^鍵No\.(\d)/); if (k) return ["k" + k[1]];
    if (msg.startsWith("鍵開け")) return ["open"];
    if (msg.startsWith("鍵閉め")) return ["close"];
    if (msg.includes("金庫チェック")) return ["safe"];
    return [];
}
function ngCellMap(r) {
    const map = {};
    ngMsgs(r).forEach(m => ngTargets(m).forEach(t => { if (!map[t]) map[t] = ngCat(m); }));
    return map;
}

// エラー一覧（上部）を種類別の色で表示
renderErrorPanel = function () {
    const el = document.getElementById('errorPanel'); if (!el) return;
    const bad = (shiftResult || []).filter(r => r && getDisplayStatus(r).status === 'NG');
    if (!shiftResult || !shiftResult.length) { el.innerHTML = ""; return; }
    if (!bad.length) { el.innerHTML = `<div class="p-4 rounded-xl border border-emerald-300 bg-emerald-50 text-emerald-800 text-sm font-bold"><i class="fa-solid fa-circle-check"></i> エラーなし：全日程で担当者が設定され、全ルールを満たしています。</div>`; return; }
    const items = bad.map(r => `<li class="flex gap-2 items-start"><span class="font-bold whitespace-nowrap text-slate-700 pt-0.5">${esc(r.date)}(${esc(r.dayOfWeek)})</span><span class="flex flex-wrap gap-1">${ngMsgs(r).map(ngChip).join('')}</span></li>`).join('');
    el.innerHTML = `<div class="p-4 rounded-xl border border-red-300 bg-red-50 space-y-2"><div class="font-bold text-red-700 text-sm"><i class="fa-solid fa-triangle-exclamation"></i> エラー・要確認：${bad.length}日（ルール設定や手動修正で調整できます）</div>${ngLegend(shiftResult)}<ul class="space-y-1.5 text-xs max-h-64 overflow-auto">${items}</ul></div>`;
};

// 割り当て結果の表：NGの欄を種類の色で塗る・詳細理由を色付きで表示
function decorateResultNG() {
    const n = orderedStaff().length;
    const col = { open: n + 1, close: n + 2, safe: n + 3, k1: [n + 4, n + 5], k2: [n + 6, n + 7], k3: [n + 8, n + 9] };
    document.querySelectorAll('#resultTableBody tr').forEach((tr, i) => {
        const r = shiftResult[i]; if (!r) return;
        tr.classList.remove('bg-red-50');
        const tds = tr.children;
        if (tds[0]) tds[0].classList.replace('bg-red-100', 'bg-white');
        if (getDisplayStatus(r).status !== 'NG') return;
        const map = ngCellMap(r);
        Object.keys(map).forEach(t => [].concat(col[t]).forEach(ci => { if (tds[ci]) tds[ci].classList.add(NG_CATS[map[t]].cell); }));
        if (tds[0]) tds[0].insertAdjacentHTML('beforeend', `<div>${ngDots(r)}</div>`);
        if (tds[n + 10]) tds[n + 10].insertAdjacentHTML('beforeend', `<div>${ngDots(r)}</div>`);
        if (tds[n + 11]) tds[n + 11].innerHTML = `<div class="flex flex-wrap gap-1 justify-center">${ngMsgs(r).map(ngChip).join('')}</div>`;
    });
}
const _renderResultP3 = renderResult;
renderResult = function () { _renderResultP3(); decorateResultNG(); };

// 手動修正の表も同じ色分け
const _renderEditTableP3 = renderEditTable;
renderEditTable = function () {
    _renderEditTableP3();
    const n = orderedStaff().length;
    const col = { open: n + 1, close: n + 2, safe: n + 3, k1: [n + 4, n + 5], k2: [n + 6, n + 7], k3: [n + 8, n + 9] };
    document.querySelectorAll('#editTableBody tr').forEach((tr, i) => {
        const r = shiftResult[i]; if (!r) return;
        tr.classList.remove('bg-red-50');
        if (getDisplayStatus(r).status !== 'NG') return;
        const tds = tr.children, map = ngCellMap(r);
        Object.keys(map).forEach(t => [].concat(col[t]).forEach(ci => { const td = tds[ci]; if (td) td.classList.add(NG_CATS[map[t]].cell); }));
        const last = tds[tds.length - 1];
        if (last) last.innerHTML = `<div class="flex flex-wrap gap-1">${ngMsgs(r).map(ngChip).join('')}</div>`;
    });
    const lg = document.getElementById('editNgLegend');
    if (lg) lg.innerHTML = (shiftResult && shiftResult.length) ? ngLegend(shiftResult) : "";
};
(function () {
    const tb = document.getElementById('editToolbar');
    if (tb) tb.insertAdjacentHTML('afterend', '<div id="editNgLegend" class="mb-3"></div>');
})();

// ===== 追加：STEP0（シフト編集）で、元のシフトを手で直せるようにする =====
let s0OnlyMembers = true;
const S0_INPUT_CLS = "w-11 px-0.5 py-1 text-center text-xs font-bold rounded border border-transparent hover:border-indigo-300 focus:border-indigo-500 focus:outline-none";
function s0IsMember(nm) {
    nm = String(nm || "");
    return staffList.some(s => [s.name, ...(s.aliases || [])].some(k => k && (nm === k || nm.includes(k))));
}
function s0Changed(r, c) {
    const o = step0State && step0State.orig;
    return !!o && String((o[r] || [])[c] ?? "") !== String((step0State.aoa[r] || [])[c] ?? "");
}
function s0ChangedCount() {
    if (!step0State || !step0State.orig) return 0;
    let n = 0;
    step0State.aoa.forEach((row, r) => { if (r >= 4) row.forEach((_, c) => { if (c >= 3 && s0Changed(r, c)) n++; }); });
    return n;
}
function s0CellCls(v, changed) {
    const s = String(v || "").trim();
    let cls = "bg-white text-slate-400";
    if (s && isWorkingShift(s)) {
        if (s.includes("超早")) cls = "bg-blue-100 text-blue-900";
        else if (s.includes("早") || isEarlyByTime(s)) cls = "bg-sky-50 text-sky-900";
        else if (isFullShift(s)) cls = "bg-violet-50 text-violet-900";
        else if (isLateShift(s)) cls = "bg-amber-50 text-amber-900";
        else cls = "bg-white text-slate-700";
    }
    return changed ? "bg-yellow-200 text-slate-900 ring-2 ring-yellow-500" : cls;
}

s0RenderPreview = function () {
    const box = document.getElementById('s0Result');
    if (!box) return;
    if (!step0State) { box.classList.add('hidden'); return; }
    box.classList.remove('hidden');
    const st = step0State;
    if (!st.orig) st.orig = deepClone(st.aoa);
    const changed = s0ChangedCount();
    document.getElementById('s0Info').textContent =
        `${st.fileName} ／ シート「${st.sheetName}」 ／ ${st.y}年${st.m}月（${st.lastDay}日分） ／ スタッフ${st.staffCount}名`;

    // 入力候補（表にある表記＋よく使う表記）
    const opts = new Set(["早", "超早", "遅", "フル", "公休", "有給", "応援", "会議", "店長会議", "研修"]);
    st.aoa.forEach((row, r) => { if (r >= 4 && row.length > 1) row.slice(3).forEach(v => { const t = s0Str(v); if (t) opts.add(t); }); });

    let html = "";
    let pendingGroup = null;
    st.aoa.forEach((row, ri) => {
        if (ri === 0 || ri === 1) return;
        if (ri <= 3) {
            html += "<tr>" + row.map((v, ci) => {
                const dow = ri === 3 ? s0Str(v) : s0Str(st.aoa[3][ci]);
                const dc = dow === "日" ? "text-red-600" : dow === "土" ? "text-blue-600" : "";
                return `<th class="px-1 py-1 whitespace-nowrap bg-slate-100 font-semibold sticky z-10 ${ri === 2 ? 'top-0' : 'top-[24px]'} ${dc}">${esc(v)}</th>`;
            }).join("") + "</tr>";
            return;
        }
        if (row.length === 1) { pendingGroup = `<tr><td colspan="${st.width}" class="px-2 py-1 text-left font-bold bg-indigo-50 text-indigo-700">${esc(row[0])}</td></tr>`; return; }
        const member = s0IsMember(row[2]);
        if (s0OnlyMembers && !member) return;
        if (pendingGroup) { html += pendingGroup; pendingGroup = null; }
        html += `<tr class="${member ? '' : 'opacity-70'}">` + row.map((v, ci) => {
            if (ci < 3) return `<td class="px-1.5 py-1 whitespace-nowrap border border-slate-100 text-left ${ci === 2 ? 'font-bold ' + (member ? 'text-indigo-700' : 'text-slate-600') : 'text-slate-500'}">${esc(v)}</td>`;
            return `<td class="p-0.5 border border-slate-100"><input value="${esc(v)}" list="s0Opts" data-r="${ri}" data-c="${ci}" onchange="s0Edit(this)" title="${esc(row[2])} ${ci - 2}日" class="${S0_INPUT_CLS} ${s0CellCls(v, s0Changed(ri, ci))}"></td>`;
        }).join("") + "</tr>";
    });
    document.getElementById('s0Preview').innerHTML = html;

    let bar = document.getElementById('s0EditBar');
    if (!bar) {
        document.getElementById('s0Preview').closest('div.overflow-auto').insertAdjacentHTML('beforebegin', '<div id="s0EditBar"></div>');
        bar = document.getElementById('s0EditBar');
    }
    bar.innerHTML = `
      <datalist id="s0Opts">${[...opts].map(o => `<option value="${esc(o)}"></option>`).join('')}</datalist>
      <div class="flex flex-wrap items-center gap-3 p-3 rounded-lg border border-amber-300 bg-amber-50 text-xs">
        <span class="font-bold text-amber-800"><i class="fa-solid fa-pen"></i> 元のシフトが間違っている場合は、下の表のマスをクリックして直接書き換えられます（候補から選択も可）。</span>
        <label class="flex items-center gap-1 cursor-pointer font-bold text-slate-700"><input type="checkbox" ${s0OnlyMembers ? 'checked' : ''} onchange="s0OnlyMembers=this.checked; s0RenderPreview();"> 登録メンバー（鍵当番の対象者）だけ表示</label>
        <span id="s0ChangedInfo" class="${changed ? 'font-bold text-amber-800' : 'text-slate-500'}">${changed ? `変更 ${changed}か所（黄色のマス）` : '変更なし'}</span>
        <button type="button" onclick="s0ResetEdits()" ${changed ? '' : 'disabled'} class="px-2.5 py-1 rounded-md border text-xs font-bold ${changed ? 'border-slate-400 bg-white hover:bg-slate-50 cursor-pointer' : 'border-slate-200 text-slate-300 cursor-not-allowed'}"><i class="fa-solid fa-rotate-left"></i> 変更を元に戻す</button>
        <span class="text-slate-500 w-full">※直した内容はすぐに「1. シフト取込」に反映され、ダウンロードする編集用Excelにも入ります（作成済みの割り当て結果は作り直しになります）。</span>
      </div>`;
};

function s0Edit(el) {
    if (!step0State) return;
    const r = Number(el.dataset.r), c = Number(el.dataset.c);
    step0State.aoa[r][c] = String(el.value || "").trim();
    el.className = S0_INPUT_CLS + " " + s0CellCls(el.value, s0Changed(r, c));
    s0FeedStep1(step0State);
    const changed = s0ChangedCount();
    const info = document.getElementById('s0ChangedInfo');
    if (info) { info.textContent = changed ? `変更 ${changed}か所（黄色のマス）` : '変更なし'; info.className = changed ? 'font-bold text-amber-800' : 'text-slate-500'; }
    const btn = document.querySelector('#s0EditBar button');
    if (btn) { btn.disabled = !changed; btn.className = "px-2.5 py-1 rounded-md border text-xs font-bold " + (changed ? "border-slate-400 bg-white hover:bg-slate-50 cursor-pointer" : "border-slate-200 text-slate-300 cursor-not-allowed"); }
    const msg = document.getElementById('s0Msg');
    if (msg) msg.textContent = (rawShiftData && rawShiftData.length)
        ? `修正をSTEP1に反映しました（${step0State.y}年${step0State.m}月・変更${changed}か所）。「1. シフト取込」へ進めます。`
        : "修正しましたが、登録メンバーが見つからず STEP1 には入りませんでした。画面2でメンバー名・「Excel上の表記」を確認してください。";
}

function s0ResetEdits() {
    if (!step0State || !step0State.orig) return;
    if (!confirm("シフト編集で直した内容をすべて元に戻します。よろしいですか？")) return;
    step0State.aoa = deepClone(step0State.orig);
    s0FeedStep1(step0State);
    s0RenderPreview();
    const msg = document.getElementById('s0Msg'); if (msg) msg.textContent = "直した内容を元に戻し、STEP1に反映しました。";
}

// ===== リセット：読込データ・割り当て結果・版・比較・集計をすべて消す =====
const _resetAllBase = resetAll;
resetAll = function () {
    const firstRules = (shiftVersions.length && shiftVersions[0].rules) ? deepClone(shiftVersions[0].rules) : null;   // 自動修正前の条件
    _resetAllBase();
    if (step0State || (rawShiftData && rawShiftData.length)) return;   // キャンセルされた
    if (firstRules) RULES = firstRules;   // 自動修正で足した条件も外す
    shiftVersions = []; activeVer = -1; fixMessage = ""; _pendingKind = null;
    cmpLeft = -1; cmpRight = -1;
    shiftResult = []; rawShiftData = [];
    staffList.forEach(s => { s.openCount = 0; s.closeCount = 0; s.safeCount = 0; });
    const set = (id, html) => { const el = document.getElementById(id); if (el) el.innerHTML = html; };
    set('statsTableBody', ""); set('fixPanel', ""); set('errorPanel', ""); set('resultTableBody', ""); set('editTableBody', "");
    set('editNgCount', ""); set('editNgLegend', ""); set('s0Preview', ""); set('s0EditBar', "");
    set('previewBody', `<tr><td colspan="${orderedStaff().length + 2}" class="px-4 py-8 text-center text-slate-400">ファイル未選択です。「0. シフト編集」または上のファイル選択からExcelを読み込んでください。</td></tr>`);
    const lbl = document.getElementById('resultVerLabel'); if (lbl) lbl.textContent = "3-1";
    const cb = document.getElementById('editNgOnly'); if (cb) cb.checked = false;
    try { renderCompare(); } catch (e) { }
    try { renderPriorityList(); } catch (e) { }
};
</script>
<style>
/* ===== 動き・見やすさ ===== */
@keyframes fadeUp { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }
@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
@keyframes popIn { from { opacity: 0; transform: scale(.92) translateY(12px); } to { opacity: 1; transform: none; } }
@keyframes ngPulse { 0%, 100% { box-shadow: inset 4px 0 0 #ef4444; } 50% { box-shadow: inset 7px 0 0 #f87171; } }
@keyframes flash { 0% { background-color: #fde68a; } 100% { background-color: transparent; } }
@keyframes toastIn { from { opacity: 0; transform: translate(-50%, 16px); } to { opacity: 1; transform: translate(-50%, 0); } }

body { background: linear-gradient(180deg, #eef2ff 0, #f8fafc 260px) !important; }
header { background: linear-gradient(120deg, #4338ca, #6366f1 55%, #8b5cf6) !important; }
nav .tab-btn { transition: color .2s, border-color .25s, background-color .2s; border-radius: 8px 8px 0 0; }
nav .tab-btn:hover { background: #eef2ff; }
.tab-content:not(.hidden) { animation: fadeUp .35s ease both; }
.tab-content:not(.hidden) > div, .tab-content:not(.hidden) > .grid > div { animation: fadeUp .45s ease both; }
.tab-content:not(.hidden) > *:nth-child(2) { animation-delay: .05s; }
.tab-content:not(.hidden) > *:nth-child(3) { animation-delay: .1s; }
.tab-content:not(.hidden) > *:nth-child(4) { animation-delay: .15s; }
.rounded-xl.shadow-sm { transition: box-shadow .25s ease, transform .25s ease; }
.rounded-xl.shadow-sm:hover { box-shadow: 0 12px 28px -14px rgba(30, 27, 75, .25); }
button { transition: transform .12s ease, box-shadow .2s ease, background-color .2s ease, color .2s ease; }
button:not(:disabled):hover { box-shadow: 0 4px 14px -4px rgba(79, 70, 229, .35); }
button:not(:disabled):active { transform: scale(.96); }
#resultTableBody tr, #editTableBody tr { animation: fadeUp .3s ease both; }
#resultTableBody tr.ng-row > td:first-child { animation: ngPulse 1.8s ease-in-out infinite; }
#resultTableBody tr.ng-row { cursor: pointer; }
#resultTableBody tr.ng-row:hover > td { filter: brightness(.97); }
.flash-row > td { animation: flash 1.6s ease; }
.pop-overlay { animation: fadeIn .2s ease both; }
.pop-card { animation: popIn .3s cubic-bezier(.2, .9, .3, 1.25) both; }
.toast { animation: toastIn .3s ease both; }
@media (prefers-reduced-motion: reduce) { *, *::before, *::after { animation: none !important; transition: none !important; } }
</style>
<script>
// ===== 追加：鍵開け優先者（既定：五家が早番の日は五家が鍵開け） =====
const _candScoreP4 = candScore;
candScore = function (name, task, date, dow, cnt) {
    let sc = _candScoreP4(name, task, date, dow, cnt);
    if (task === 'open' && RULES.openPreferOn && name === RULES.openPrefer) {
        const sh = (_shiftIdx.byDate[date] || {})[name];
        if (canOpenShift(sh || "公休")) sc -= 1000;
    }
    return sc;
};

const _renameStaffP4 = renameStaff;
renameStaff = function (i, value) {
    const old = staffList[i] ? staffList[i].name : null;
    _renameStaffP4(i, value);
    if (old && RULES.openPrefer === old && staffList[i] && staffList[i].name !== old) { RULES.openPrefer = staffList[i].name; renderRulesPanel(); }
};

const _renderRulesPanelP4 = renderRulesPanel;
renderRulesPanel = function () {
    _renderRulesPanelP4();
    const el = document.getElementById('rulesPanel'); if (!el) return;
    if (RULES.openPreferOn && !staffList.some(s => s.name === RULES.openPrefer)) RULES.openPrefer = staffList[0] ? staffList[0].name : "";
    const ic = "border border-slate-300 rounded px-2 py-1 text-xs bg-white";
    el.insertAdjacentHTML('afterbegin', `
      <div class="bg-white p-5 rounded-xl shadow-sm border-2 border-indigo-300 space-y-3 text-xs text-slate-600">
        <h3 class="font-bold text-slate-700 text-sm flex items-center gap-2"><i class="fa-solid fa-vault text-amber-600"></i>金庫チェックの担当条件</h3>
        <label class="flex items-start gap-1.5 cursor-pointer"><input type="checkbox" class="mt-0.5" ${RULES.safeLateOnly ? 'checked' : ''} onchange="setRule('safeLateOnly', this.checked)"><span><b>遅番以外のシフトの人は金庫チェックができない</b>（早番・超早の人、会議・店長会議・研修・応援・休みなど店舗にいない人には割り当てない。遅番がいない日は担当なしでNG表示）</span></label>
        <label class="flex items-start gap-1.5 cursor-pointer ml-5"><input type="checkbox" class="mt-0.5" ${RULES.safeFullOk !== false ? 'checked' : ''} onchange="setRule('safeFullOk', this.checked)"><span><b>「フル」の人は金庫チェックできる</b>（遅番として扱う）</span></label>
        <hr class="border-slate-200">
        <h3 class="font-bold text-slate-700 text-sm flex items-center gap-2"><i class="fa-solid fa-door-open text-indigo-600"></i>鍵開けの優先担当</h3>
        <label class="flex items-start gap-1.5 cursor-pointer"><input type="checkbox" class="mt-0.5" ${RULES.openPreferOn ? 'checked' : ''} onchange="setRule('openPreferOn', this.checked)">
          <span><select onchange="setRule('openPrefer', this.value)" class="${ic} font-bold">${staffList.map(s => `<option value="${esc(s.name)}" ${s.name === RULES.openPrefer ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select>
          が早番（超早・早・フル）の日は、優先的にこの人が鍵開けをする（前日の夜までに鍵が渡るように鍵の持ち回しも調整）</span></label>
        <label class="flex items-center gap-1.5">前日に鍵を渡す強さ<input type="number" min="0" value="${Number(RULES.openPreferWeight) || 60}" onchange="setRule('openPreferWeight', this.value, 1)" class="w-16 ${ic}"><span class="text-slate-400">（大きいほど受け渡しを増やしてでも優先者に鍵を渡す）</span></label>
      </div>`);
};
(function () {
    const li = document.querySelector('#tab2 ul.list-disc');
    if (li) li.insertAdjacentHTML('afterbegin', '<li><strong>鍵開けの優先担当</strong>：五家（変更可）が早番の日は、五家が優先的に鍵開けをします。そのため前日の夜までに五家へ鍵が渡るように持ち回しを調整します。</li>');
    try { renderRulesPanel(); } catch (e) { }
})();

// ===== 追加：NGのポップアップ表示 =====
function closePop() { const el = document.getElementById('ngPop'); if (el) el.remove(); }
function openPop(html, wide) {
    closePop();
    document.body.insertAdjacentHTML('beforeend', `
      <div id="ngPop" class="pop-overlay fixed inset-0 z-[60] bg-slate-900/45 backdrop-blur-[2px] flex items-center justify-center p-4" onclick="if(event.target===this) closePop()">
        <div class="pop-card bg-white rounded-2xl shadow-2xl w-full ${wide ? 'max-w-3xl' : 'max-w-lg'} max-h-[85vh] flex flex-col overflow-hidden">${html}</div>
      </div>`);
}
document.addEventListener('keydown', e => { if (e.key === 'Escape') { closePop(); closeDayPop(); } });

function toast(msg, ok) {
    const old = document.getElementById('toast'); if (old) old.remove();
    document.body.insertAdjacentHTML('beforeend', `<div id="toast" class="toast fixed bottom-6 left-1/2 z-[70] px-5 py-3 rounded-xl shadow-xl text-sm font-bold text-white ${ok ? 'bg-emerald-600' : 'bg-slate-800'}" style="transform:translateX(-50%)">${esc(msg)}</div>`);
    setTimeout(() => { const t = document.getElementById('toast'); if (t) { t.style.transition = 'opacity .4s'; t.style.opacity = '0'; setTimeout(() => t.remove(), 450); } }, 2600);
}

function showNgSummaryPop(extraMsg) {
    const v = shiftVersions[activeVer];
    const bad = (shiftResult || []).filter(r => r && getDisplayStatus(r).status === 'NG');
    const label = v ? `${v.label} 割り当て結果` : '割り当て結果';
    if (!bad.length) {
        openPop(`
          <div class="p-8 text-center space-y-3">
            <div class="w-16 h-16 mx-auto rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center text-3xl"><i class="fa-solid fa-circle-check"></i></div>
            <div class="text-lg font-bold text-slate-800">NGはありません</div>
            <div class="text-sm text-slate-500">「${esc(label)}」は全日程で担当者が決まり、すべてのルールを満たしています。</div>
            ${extraMsg ? `<div class="text-xs text-emerald-700 font-bold">${esc(extraMsg)}</div>` : ''}
            <button type="button" onclick="closePop()" class="mt-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-6 py-2 rounded-lg cursor-pointer">OK</button>
          </div>`);
        return;
    }
    const items = bad.map(r => {
        const i = shiftResult.indexOf(r);
        return `<li class="p-2 rounded-lg border border-slate-200 hover:border-indigo-300 hover:bg-indigo-50/40 cursor-pointer" onclick="closePop(); jumpToResultRow(${i})">
          <div class="flex items-center gap-2 mb-1"><span class="font-bold text-slate-800">${esc(r.date)}(${esc(r.dayOfWeek)})</span>${ngDots(r)}<span class="ml-auto text-[10px] text-indigo-500">表で見る <i class="fa-solid fa-arrow-right"></i></span></div>
          <div class="flex flex-wrap gap-1">${ngMsgs(r).map(ngChip).join('')}</div></li>`;
    }).join('');
    openPop(`
      <div class="px-5 py-4 bg-gradient-to-r from-red-500 to-rose-500 text-white flex items-center gap-3">
        <i class="fa-solid fa-triangle-exclamation text-2xl"></i>
        <div><div class="font-bold text-base">NGが ${bad.length}日 あります</div><div class="text-xs opacity-90">${esc(label)}</div></div>
        <button type="button" onclick="closePop()" class="ml-auto w-8 h-8 rounded-full hover:bg-white/20 cursor-pointer"><i class="fa-solid fa-xmark"></i></button>
      </div>
      <div class="p-4 space-y-3 overflow-auto">
        ${extraMsg ? `<div class="text-xs font-bold text-amber-800 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">${esc(extraMsg)}</div>` : ''}
        ${ngLegend(shiftResult)}
        <ul class="space-y-2 text-xs">${items}</ul>
      </div>
      <div class="px-4 py-3 border-t border-slate-200 bg-slate-50 flex flex-wrap gap-2 justify-end">
        <button type="button" onclick="closePop(); autoFix();" class="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-4 py-2 rounded-lg text-sm cursor-pointer"><i class="fa-solid fa-wand-magic-sparkles"></i> 自動で修正する</button>
        <button type="button" onclick="closePop(); goManualFix();" class="bg-white border-2 border-amber-500 text-amber-700 hover:bg-amber-50 font-semibold px-4 py-2 rounded-lg text-sm cursor-pointer"><i class="fa-solid fa-hand"></i> 手動で一気に直す</button>
        <button type="button" onclick="closePop()" class="bg-white border border-slate-300 text-slate-600 hover:bg-slate-100 font-semibold px-4 py-2 rounded-lg text-sm cursor-pointer">閉じる</button>
      </div>`, true);
}

// NGの行をクリック → その日の内容を小さなポップで表示
function closeDayPop() { const el = document.getElementById('dayPop'); if (el) el.remove(); }
function showDayPop(i, ev) {
    closeDayPop();
    const r = shiftResult[i]; if (!r) return;
    const shifts = orderedStaff().map(s => `<span class="inline-flex items-center gap-1 text-[11px]"><b>${esc(s.name)}</b>${shiftBadge((r.shiftsByStaff || {})[s.name] || "公休")}</span>`).join('');
    const kh = r.keyHolders || {};
    document.body.insertAdjacentHTML('beforeend', `
      <div id="dayPop" class="pop-card fixed z-[65] w-[22rem] max-w-[92vw] bg-white rounded-xl shadow-2xl border border-slate-200 overflow-hidden">
        <div class="px-3 py-2 bg-red-50 border-b border-red-200 flex items-center gap-2"><span class="font-bold text-red-700 text-sm">${esc(r.date)}(${esc(r.dayOfWeek)}) のNG</span>${ngDots(r)}
          <button type="button" onclick="closeDayPop()" class="ml-auto w-6 h-6 rounded-full hover:bg-red-100 text-red-600 cursor-pointer"><i class="fa-solid fa-xmark"></i></button></div>
        <div class="p-3 space-y-2 text-xs">
          <div class="flex flex-wrap gap-1">${ngMsgs(r).map(ngChip).join('')}</div>
          <div class="flex flex-wrap gap-x-3 gap-y-1">${shifts}</div>
          <div class="text-slate-600">鍵開け <b>${esc(r.openRaw)}</b> ／ 鍵閉め <b>${esc(r.closeRaw)}</b> ／ 金庫 <b>${esc(r.safeRaw)}</b></div>
          <div class="text-slate-600">退勤時の鍵：<span class="text-red-600 font-bold">1</span> ${esc(kh["鍵No.1"])}　<span class="text-amber-600 font-bold">2</span> ${esc(kh["鍵No.2"])}　<span class="text-blue-600 font-bold">3</span> ${esc(kh["鍵No.3"])}</div>
          <div class="flex gap-2 pt-1">
            <button type="button" onclick="closeDayPop(); goManualFixDay(${i});" class="flex-1 bg-amber-500 hover:bg-amber-600 text-white font-bold px-3 py-1.5 rounded-lg cursor-pointer"><i class="fa-solid fa-hand"></i> この日を手動で直す</button>
            <button type="button" onclick="closeDayPop(); autoFix();" class="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-3 py-1.5 rounded-lg cursor-pointer"><i class="fa-solid fa-wand-magic-sparkles"></i> 自動で修正</button>
          </div>
        </div>
      </div>`);
    const el = document.getElementById('dayPop');
    const w = el.offsetWidth, h = el.offsetHeight;
    let x = (ev ? ev.clientX : innerWidth / 2) + 12, y = (ev ? ev.clientY : innerHeight / 2) + 12;
    if (x + w > innerWidth - 8) x = Math.max(8, innerWidth - w - 8);
    if (y + h > innerHeight - 8) y = Math.max(8, (ev ? ev.clientY : innerHeight / 2) - h - 12);
    el.style.left = x + 'px'; el.style.top = y + 'px';
}
document.addEventListener('click', e => {
    const p = document.getElementById('dayPop');
    if (p && !p.contains(e.target) && !e.target.closest('tr.ng-row')) closeDayPop();
});

function jumpToResultRow(i) {
    switchTab('tab3');
    const tr = document.querySelectorAll('#resultTableBody tr')[i];
    if (tr) { tr.scrollIntoView({ behavior: 'smooth', block: 'center' }); tr.classList.remove('flash-row'); void tr.offsetWidth; tr.classList.add('flash-row'); }
}
function goManualFixDay(i) {
    goManualFix();
    setTimeout(() => {
        const tr = document.querySelectorAll('#editTableBody tr')[i];
        if (tr) { tr.scrollIntoView({ behavior: 'smooth', block: 'center' }); tr.classList.add('flash-row'); }
    }, 60);
}

// 新しい版ができたら（作成・自動修正・再計算）ポップアップで知らせる
const _pushVersionP4 = pushVersion;
pushVersion = function (kind, changes, ngBefore) { _pushVersionP4(kind, changes, ngBefore); const v = shiftVersions[activeVer]; if (v) v.popPending = true; };

let _lastFixMsg = "";
const _renderResultP4 = renderResult;
renderResult = function () {
    _renderResultP4();
    // 行の動き（順番にふわっと表示）とNG行のクリック
    document.querySelectorAll('#resultTableBody tr').forEach((tr, i) => {
        tr.style.animationDelay = Math.min(i * 18, 500) + 'ms';
        const r = shiftResult[i];
        if (r && getDisplayStatus(r).status === 'NG') {
            tr.classList.add('ng-row');
            tr.title = "クリックでこの日のNGを表示";
            tr.addEventListener('click', ev => { if (ev.target.closest('button,select,input,a')) return; ev.stopPropagation(); showDayPop(i, ev); });
        }
    });
    const tab3 = document.getElementById('tab3');
    const visible = tab3 && !tab3.classList.contains('hidden');
    const v = shiftVersions[activeVer];
    if (!visible || !shiftResult || !shiftResult.length) return;
    if (v && v.popPending) {
        v.popPending = false;
        _lastFixMsg = fixMessage;
        const ngDays = ngScore(shiftResult).days;
        if (v.kind.startsWith("自動修正")) {
            if (ngDays === 0) showNgSummaryPop(`${v.label}：NG ${v.ngBefore}日 → 0日 に修正しました`);
            else showNgSummaryPop(fixMessage.replace(/^⚠\s*/, ""));
        } else showNgSummaryPop();
    } else if (fixMessage && fixMessage !== _lastFixMsg) {
        _lastFixMsg = fixMessage;
        if (fixMessage.startsWith('⚠')) showNgSummaryPop(fixMessage.replace(/^⚠\s*/, ""));
    }
};
// 版の切り替え時は小さなお知らせだけ
const _selectVersionP4 = selectVersion;
selectVersion = function (i) { _selectVersionP4(i); const v = shiftVersions[activeVer]; if (v) toast(`${v.label} 割り当て結果を表示中（NG ${ngScore(shiftResult).days}日）`, ngScore(shiftResult).days === 0); };

const _renderEditTableP4 = renderEditTable;
renderEditTable = function () {
    _renderEditTableP4();
    document.querySelectorAll('#editTableBody tr').forEach((tr, i) => { tr.style.animationDelay = Math.min(i * 12, 350) + 'ms'; });
};
</script>
<script>
const _autoFixP4 = autoFix;
autoFix = function () { _lastFixMsg = ""; _autoFixP4(); };
</script>
<style>
/* シフト入力マス：クリックすると広がって自由に入力しやすく */
.shift-in { transition: width .15s ease, box-shadow .15s ease; }
.shift-in:focus { width: 5.5rem !important; position: relative; z-index: 5; box-shadow: 0 0 0 3px rgba(99, 102, 241, .35); background: #fff !important; color: #0f172a !important; }
#ocrDrop.drag { border-color: #6366f1; background: #eef2ff; }
</style>
<script>
// ===== 追加：シフトの手入力（自由に文字を打てる）を STEP0・STEP1 共通で =====
const SHIFT_SUGGEST = ["早", "超早", "遅", "フル", "公休", "有給", "応援", "会議", "店長会議", "研修", "グランド", "方針", "欠勤"];

function shiftInputCls(v, changed) {
    return "shift-in w-12 px-0.5 py-1 text-center text-xs font-bold rounded border border-slate-200 hover:border-indigo-400 focus:outline-none " + s0CellCls(v, changed);
}
// Enterで右隣のマスへ移動（続けて打ちやすく）
document.addEventListener('keydown', e => {
    const el = e.target;
    if (e.key !== 'Enter' || !el.classList || !el.classList.contains('shift-in')) return;
    e.preventDefault();
    const all = [...document.querySelectorAll('input.shift-in')].filter(x => x.offsetParent !== null);
    const i = all.indexOf(el);
    el.blur();
    const nx = all[i + 1];
    if (nx) { nx.focus(); nx.select(); }
});
document.addEventListener('focusin', e => { if (e.target.classList && e.target.classList.contains('shift-in')) e.target.select(); });

// 1人・1日のシフトを書き換える（STEP0の表があればそちらを直してSTEP1へ流す）
function s0RowOfStaff(name) {
    if (!step0State) return -1;
    const st = staffList.find(s => s.name === name); if (!st) return -1;
    const keys = [st.name, ...(st.aliases || [])].filter(Boolean);
    let hit = -1;
    step0State.aoa.forEach((row, r) => {
        if (r < 4 || row.length <= 1 || hit >= 0) return;
        const nm = s0Str(row[2]);
        if (keys.some(k => nm === k)) hit = r;
    });
    if (hit < 0) step0State.aoa.forEach((row, r) => { if (r >= 4 && row.length > 1 && hit < 0 && keys.some(k => s0Str(row[2]).includes(k))) hit = r; });
    return hit;
}
function applyShiftFix(name, date, value, silent) {
    value = String(value || "").trim();
    const d = parseInt(String(date).split("/").pop(), 10);
    const r = s0RowOfStaff(name);
    if (step0State && r >= 0) {
        step0State.aoa[r][2 + d] = value;
        s0FeedStep1(step0State);
        s0RenderPreview();
    } else {
        const it = (rawShiftData || []).find(x => x.name === name && x.date === date);
        if (it) it.shiftType = value || "公休";
        else if (rawShiftData && rawShiftData.length) {
            const dow = (rawShiftData.find(x => x.date === date) || {}).dayOfWeek || "";
            rawShiftData.push({ date, dayOfWeek: dow, name, shiftType: value || "公休" });
        }
        shiftResult = [];
        renderPreview();
    }
    if (!silent) toast(`${name} の ${date} を「${value || "公休"}」に直しました`, true);
}

// STEP1の一覧表も、マスに直接入力できるようにする
renderPreview = function () {
    renderStaffHeaders();
    const tbody = document.getElementById('previewBody'); if (!tbody) return;
    const staff = orderedStaff();
    if (!rawShiftData || !rawShiftData.length) {
        tbody.innerHTML = `<tr><td colspan="${staff.length + 2}" class="px-4 py-8 text-center text-slate-400">ファイル未選択です。「0. シフト編集」または上のファイル選択からExcelを読み込んでください。</td></tr>`;
        return;
    }
    const dateMap = {}, order = [];
    rawShiftData.forEach(it => {
        if (!dateMap[it.date]) { dateMap[it.date] = { dow: it.dayOfWeek, s: {} }; order.push(it.date); }
        dateMap[it.date].s[it.name] = it.shiftType;
    });
    tbody.innerHTML = `<tr><td colspan="${staff.length + 2}" class="px-3 py-1.5 text-left text-[11px] text-amber-800 bg-amber-50"><i class="fa-solid fa-pen"></i> マスをクリックすると自由に文字を入力できます（例：早、遅、9 1830）。Enterで右のマスへ移動します。</td></tr>` +
        order.map(dt => {
            const o = dateMap[dt];
            const dc = o.dow === "日" ? "text-red-600" : o.dow === "土" ? "text-blue-600" : "";
            return `<tr><td class="px-3 py-1.5 font-medium whitespace-nowrap ${dc}">${esc(dt)}</td><td class="px-3 py-1.5 ${dc}">${esc(o.dow)}</td>` +
                staff.map(s => { const v = o.s[s.name] || "公休"; return `<td class="p-0.5"><input value="${esc(v)}" list="shiftSuggest" class="${shiftInputCls(v)}" title="${esc(s.name)} ${esc(dt)}" onchange="applyShiftFix('${esc(s.name)}','${esc(dt)}',this.value)"></td>`; }).join('') + `</tr>`;
        }).join('');
};
document.body.insertAdjacentHTML('beforeend', `<datalist id="shiftSuggest">${SHIFT_SUGGEST.map(v => `<option value="${v}"></option>`).join('')}</datalist>`);

// STEP0の表：入力マスを「自由入力しやすい」見た目に差し替え
const _s0RenderPreviewP5 = s0RenderPreview;
s0RenderPreview = function () {
    _s0RenderPreviewP5();
    document.querySelectorAll('#s0Preview input').forEach(inp => {
        const r = Number(inp.dataset.r), c = Number(inp.dataset.c);
        inp.className = shiftInputCls(inp.value, s0Changed(r, c));
    });
    const bar = document.querySelector('#s0EditBar .font-bold.text-amber-800');
    if (bar) bar.innerHTML = '<i class="fa-solid fa-pen"></i> シフトが間違っているマスは、クリックして<u>自由に文字を入力</u>できます（例：早、遅、フル、9 1830）。候補から選ぶことも可能。Enterで右のマスへ移動。';
};
s0Edit = (function (base) {
    return function (el) { base(el); const r = Number(el.dataset.r), c = Number(el.dataset.c); el.className = shiftInputCls(el.value, s0Changed(r, c)); };
})(s0Edit);

// ===== 追加：画像認識（OCR）でシフトを二重チェック =====
let ocrWorker = null, ocrImage = null, ocrState = null;

function normShift(v) {
    let s = toHalfWidth(String(v === undefined || v === null ? "" : v)).replace(/[\s　|｜\[\]()（）・.,、。_\-ー―]/g, "").replace(/ﾌﾙ/g, "フル");
    if (!s) return null;
    if (/^\d/.test(s)) {
        // 「9 1830」「0900-1830」「91830」「101900」などの開始時刻
        const raw = toHalfWidth(String(v)).trim();
        const sp = raw.match(/^(\d{1,2})[\s:：時]/);
        let h = null;
        if (sp) h = +sp[1];
        else { const tok = (s.match(/^\d+/) || [""])[0]; h = tok.length <= 2 ? +tok : tok.length <= 4 ? Math.floor(+tok / 100) : tok.length === 5 ? +tok[0] : +tok.slice(0, 2); }
        return (h !== null && h >= 0 && h <= 24) ? "T" + h : null;
    }
    if (s.includes("超")) return "超早";
    if (s.includes("早")) return "早";
    if (s.includes("遅")) return "遅";
    if (/フ|ル/.test(s)) return "フル";
    if (s.includes("有")) return "有給";
    if (s.includes("応")) return "応援";
    if (/会|議|MTG|ミーティング/i.test(s)) return "会議";
    if (s.includes("研")) return "研修";
    if (s.includes("グ")) return "グランド";
    if (s.includes("方")) return "方針";
    if (s.includes("欠")) return "欠勤";
    if (/公|休/.test(s)) return "公休";
    return null;
}
const normData = v => normShift(v) || (String(v || "").trim() ? "?" : "公休");
const normLabel = (k, raw) => k && k.startsWith("T") ? toHalfWidth(String(raw)).replace(/\s+/g, " ").trim() : k;

function loadTesseract() {
    if (window.Tesseract) return Promise.resolve();
    return new Promise((ok, ng) => {
        const s = document.createElement('script');
        s.src = "https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js";
        s.onload = () => ok(); s.onerror = () => ng(new Error("画像認識の部品を読み込めませんでした（インターネット接続を確認してください）"));
        document.head.appendChild(s);
    });
}

function ocrSetImage(file) {
    if (!file || !file.type || !file.type.startsWith("image/")) { toast("画像ファイル（スクリーンショットなど）を選んでください"); return; }
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => { ocrImage = img; ocrState = null; renderOcrCard(); toast("画像を読み込みました。「読み取って照合する」を押してください", true); };
    img.src = url;
}

// 画像を拡大・白黒にして読み取りやすくする
function ocrPrepare(img) {
    const scale = Math.max(1, Math.min(3, 2400 / img.naturalWidth));
    const cv = document.createElement('canvas');
    cv.width = Math.round(img.naturalWidth * scale); cv.height = Math.round(img.naturalHeight * scale);
    const ctx = cv.getContext('2d');
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(img, 0, 0, cv.width, cv.height);
    const d = ctx.getImageData(0, 0, cv.width, cv.height), p = d.data;
    for (let i = 0; i < p.length; i += 4) {
        let g = 0.299 * p[i] + 0.587 * p[i + 1] + 0.114 * p[i + 2];
        g = g < 165 ? Math.max(0, g * 0.6) : 255;                 // 色付きのマス背景を飛ばして文字だけ残す
        p[i] = p[i + 1] = p[i + 2] = g;
    }
    ctx.putImageData(d, 0, 0);
    return cv;
}

function ocrWordsOf(data) {
    let ws = data.words;
    if ((!ws || !ws.length) && data.blocks) {
        ws = [];
        data.blocks.forEach(b => (b.paragraphs || []).forEach(pg => (pg.lines || []).forEach(l => (l.words || []).forEach(w => ws.push(w)))));
    }
    return (ws || []).filter(w => w && w.text && w.text.trim() && w.bbox).map(w => {
        const b = w.bbox;
        return { t: w.text.trim(), x0: b.x0, x1: b.x1, y0: b.y0, y1: b.y1, cx: (b.x0 + b.x1) / 2, cy: (b.y0 + b.y1) / 2, h: b.y1 - b.y0, conf: w.confidence };
    });
}

// 読み取った文字を「行」と「日付の列」に並べ直す
function ocrAnalyze(words) {
    if (!words.length) return { error: "文字を読み取れませんでした。画像がはっきり写っているか確認してください。" };
    const hs = words.map(w => w.h).sort((a, b) => a - b), mh = hs[Math.floor(hs.length / 2)] || 10;
    const sorted = words.slice().sort((a, b) => a.cy - b.cy);
    const rows = [];
    sorted.forEach(w => {
        const row = rows.find(r => Math.abs(r.cy - w.cy) < mh * 0.6);
        if (row) { row.w.push(w); row.cy = row.w.reduce((t, x) => t + x.cy, 0) / row.w.length; }
        else rows.push({ cy: w.cy, w: [w] });
    });
    rows.forEach(r => r.w.sort((a, b) => a.cx - b.cx));

    // 日付の行（1, 2, 3 … が横に並ぶ行）
    const dayOf = t => {
        const s = toHalfWidth(t).replace(/[日\s]/g, "");
        let m = s.match(/^\d{1,2}\/(\d{1,2})$/); if (m) return +m[1];
        m = s.match(/^(\d{1,2})$/); if (m && +m[1] >= 1 && +m[1] <= 31) return +m[1];
        return null;
    };
    let header = null;
    rows.forEach(r => {
        const ds = r.w.map(w => ({ d: dayOf(w.t), x: w.cx })).filter(o => o.d !== null);
        let inc = 0; for (let i = 1; i < ds.length; i++) if (ds[i].d > ds[i - 1].d) inc++;
        if (ds.length >= 3 && inc >= (ds.length - 1) * 0.7 && (!header || ds.length > header.ds.length)) header = { row: r, ds };
    });
    if (!header) return { error: "日付の行（1, 2, 3 …）が読み取れませんでした。日付の行も写るように画像を貼ってください。" };
    const dayX = {}; header.ds.forEach(o => { if (dayX[o.d] === undefined) dayX[o.d] = o.x; });
    const xs = Object.values(dayX).sort((a, b) => a - b);
    const gaps = []; for (let i = 1; i < xs.length; i++) gaps.push(xs[i] - xs[i - 1]);
    gaps.sort((a, b) => a - b);
    const colW = gaps.length ? gaps[Math.floor(gaps.length / 2)] : mh * 3;
    const dayEntries = Object.entries(dayX).map(([d, x]) => ({ d: +d, x }));
    const leftEdge = Math.min(...xs) - colW * 0.6;

    const out = [];
    rows.forEach(r => {
        if (r === header.row || r.cy < header.row.cy) return;
        const nameText = r.w.filter(w => w.cx < leftEdge).map(w => w.t).join("");
        const cells = {};
        r.w.filter(w => w.cx >= leftEdge).forEach(w => {
            let best = null;
            dayEntries.forEach(e => { const dist = Math.abs(e.x - w.cx); if (dist < colW * 0.6 && (!best || dist < best.dist)) best = { d: e.d, dist }; });
            if (best) (cells[best.d] = cells[best.d] || []).push(w);
        });
        const cellText = {};
        Object.keys(cells).forEach(d => { cellText[d] = cells[d].sort((a, b) => a.cx - b.cx).map(w => w.t).join(""); });
        const shiftLike = Object.values(cellText).filter(t => normShift(t)).length;
        if (shiftLike >= 2) out.push({ nameText, cells: cellText, staff: "" });
    });
    // 名前で自動対応（氏名・Excel上の表記で一致、または名前の文字が2文字以上含まれる）
    const used = new Set();
    out.forEach(row => {
        const nt = row.nameText.replace(/\s/g, "");
        let hit = staffList.find(s => !used.has(s.name) && [s.name, ...(s.aliases || [])].some(k => k && nt.includes(k)));
        if (!hit) hit = staffList.find(s => !used.has(s.name) && [s.name, ...(s.aliases || [])].some(k => k && k.length >= 2 && [...new Set(k)].filter(ch => nt.includes(ch)).length >= 2));
        if (hit) { row.staff = hit.name; used.add(hit.name); }
    });
    return { rows: out, days: Object.keys(dayX).map(Number).sort((a, b) => a - b) };
}

async function runOcr() {
    if (!ocrImage) { toast("先に画像を貼り付けてください"); return; }
    if (!rawShiftData || !rawShiftData.length) { toast("先にシフト表（Excel）を読み込んでください"); return; }
    const prog = document.getElementById('ocrProg');
    const setProg = (t, v) => { if (prog) prog.innerHTML = `<div class="text-xs text-slate-600 mb-1">${esc(t)}</div><div class="h-2 rounded bg-slate-200 overflow-hidden"><div class="h-full bg-indigo-500 transition-all" style="width:${Math.round((v || 0) * 100)}%"></div></div>`; };
    try {
        setProg("画像認識の準備中…（初回は日本語データの読み込みに少し時間がかかります）", 0.05);
        await loadTesseract();
        if (!ocrWorker) {
            ocrWorker = await Tesseract.createWorker('jpn', 1, {
                logger: m => { if (m.status) setProg(m.status === 'recognizing text' ? "文字を読み取り中…" : "準備中…（" + m.status + "）", m.progress); }
            });
            await ocrWorker.setParameters({ tessedit_pageseg_mode: '11', preserve_interword_spaces: '1' });
        }
        setProg("文字を読み取り中…", 0.1);
        const { data } = await ocrWorker.recognize(ocrPrepare(ocrImage));
        const res = ocrAnalyze(ocrWordsOf(data));
        res.rawText = data.text || "";
        ocrState = res;
        if (prog) prog.innerHTML = "";
        renderOcrCard();
        if (res.error) toast(res.error);
    } catch (err) {
        console.error(err);
        if (prog) prog.innerHTML = `<div class="text-xs font-bold text-red-600">${esc(err.message || "画像認識でエラーが発生しました")}</div>`;
    }
}

function ocrCompare() {
    if (!ocrState || !ocrState.rows) return null;
    const dataOf = (n, d) => { const it = (rawShiftData || []).find(x => x.name === n && parseInt(String(x.date).split("/").pop(), 10) === d); return it ? { v: it.shiftType, date: it.date } : null; };
    const res = { match: 0, diff: 0, unread: 0, rows: [] };
    ocrState.rows.forEach((row, ri) => {
        if (!row.staff) return;
        const cells = ocrState.days.map(d => {
            const dv = dataOf(row.staff, d);
            if (!dv) return { d, state: "none" };
            const raw = row.cells[d];
            const k = raw ? normShift(raw) : null;
            if (!raw || !k) { res.unread++; return { d, date: dv.date, data: dv.v, raw: raw || "", state: "unread" }; }
            if (k === normData(dv.v)) { res.match++; return { d, date: dv.date, data: dv.v, raw, state: "match" }; }
            res.diff++;
            return { d, date: dv.date, data: dv.v, raw, img: normLabel(k, raw), state: "diff" };
        });
        res.rows.push({ ri, staff: row.staff, cells });
    });
    return res;
}

function ocrSetStaff(ri, v) { if (ocrState && ocrState.rows[ri]) { ocrState.rows[ri].staff = v; renderOcrCard(); } }
function ocrApply(staff, date, v) { applyShiftFix(staff, date, v); renderOcrCard(); }
function ocrApplyAll() {
    const c = ocrCompare(); if (!c || !c.diff) return;
    if (!confirm(`画像と食い違う ${c.diff}か所 を、画像で読み取った値に直します。\n（画像認識は読み間違えることもあるので、念のため内容を確認してください）\nよろしいですか？`)) return;
    c.rows.forEach(r => r.cells.forEach(x => { if (x.state === "diff") applyShiftFix(r.staff, x.date, x.img, true); }));
    toast(`${c.diff}か所を画像の値に直しました`, true);
    renderOcrCard();
}
function ocrClear() { ocrImage = null; ocrState = null; renderOcrCard(); }

function renderOcrCard() {
    const el = document.getElementById('ocrCard'); if (!el) return;
    const hasData = rawShiftData && rawShiftData.length;
    const cmp = ocrCompare();
    let result = "";
    if (ocrState && ocrState.error) result = `<div class="p-3 rounded-lg bg-red-50 border border-red-200 text-xs font-bold text-red-700"><i class="fa-solid fa-circle-exclamation"></i> ${esc(ocrState.error)}</div>`;
    else if (ocrState) {
        const mapRows = ocrState.rows.map((r, i) => `<div class="flex flex-wrap items-center gap-2 text-xs"><span class="text-slate-400">行${i + 1}</span><span class="px-1.5 py-0.5 rounded bg-slate-100 font-mono max-w-[14rem] truncate" title="${esc(r.nameText)}">${esc(r.nameText || "（名前読めず）")}</span><i class="fa-solid fa-arrow-right text-slate-300"></i>
            <select onchange="ocrSetStaff(${i}, this.value)" class="border border-slate-300 rounded px-1.5 py-0.5 text-xs ${r.staff ? 'font-bold text-indigo-700' : 'text-slate-400'}"><option value="">（照合しない）</option>${staffList.map(s => `<option value="${esc(s.name)}" ${s.name === r.staff ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select>
            <span class="text-slate-400">${Object.keys(r.cells).length}マス読取</span></div>`).join('');
        const table = cmp && cmp.rows.length ? `
          <div class="overflow-auto border border-slate-200 rounded-lg max-h-[50vh]">
            <table class="compact-table text-xs text-center border-separate border-spacing-0">
              <thead class="sticky top-0 bg-slate-100"><tr><th class="px-2 py-1 sticky left-0 bg-slate-100">氏名</th>${ocrState.days.map(d => `<th class="px-1 py-1 min-w-[3.2rem]">${d}日</th>`).join('')}</tr></thead>
              <tbody>${cmp.rows.map(r => `<tr><td class="px-2 py-1 font-bold sticky left-0 bg-white whitespace-nowrap">${esc(r.staff)}</td>${r.cells.map(x => {
                if (x.state === "none") return `<td class="px-1 py-1 text-slate-300">-</td>`;
                if (x.state === "match") return `<td class="px-1 py-1 bg-emerald-50 text-emerald-800 font-bold border border-white" title="画像：${esc(x.raw)}">${esc(x.data)} <i class="fa-solid fa-check text-[9px]"></i></td>`;
                if (x.state === "unread") return `<td class="px-1 py-1 bg-slate-50 text-slate-400 border border-white" title="画像から読み取れませんでした${x.raw ? '（' + esc(x.raw) + '）' : ''}">${esc(x.data)}<div class="text-[9px]">読取不可</div></td>`;
                return `<td class="px-1 py-1 bg-red-100 border-2 border-red-400 text-red-800" title="データ：${esc(x.data)} ／ 画像：${esc(x.raw)}"><div class="font-bold line-through decoration-red-400">${esc(x.data)}</div><div class="text-[10px]">画像:${esc(x.img)}</div>
                    <button type="button" onclick="ocrApply('${esc(r.staff)}','${esc(x.date)}','${esc(x.img)}')" class="mt-0.5 px-1 rounded bg-red-500 hover:bg-red-600 text-white text-[10px] cursor-pointer">画像の値に</button></td>`;
              }).join('')}</tr>`).join('')}</tbody>
            </table>
          </div>` : `<div class="text-xs text-slate-500">照合する人を左の一覧で選んでください。</div>`;
        result = `
          <div class="grid grid-cols-3 gap-2 text-center">
            <div class="p-2 rounded-lg bg-emerald-50 border border-emerald-200"><div class="text-[11px] text-emerald-700">一致</div><div class="text-lg font-bold text-emerald-700">${cmp ? cmp.match : 0}</div></div>
            <div class="p-2 rounded-lg bg-red-50 border border-red-200"><div class="text-[11px] text-red-700">食い違い</div><div class="text-lg font-bold text-red-700">${cmp ? cmp.diff : 0}</div></div>
            <div class="p-2 rounded-lg bg-slate-50 border border-slate-200"><div class="text-[11px] text-slate-500">読み取れず</div><div class="text-lg font-bold text-slate-500">${cmp ? cmp.unread : 0}</div></div>
          </div>
          ${cmp && cmp.diff ? `<div class="flex flex-wrap items-center gap-2"><button type="button" onclick="ocrApplyAll()" class="bg-red-500 hover:bg-red-600 text-white text-xs font-bold px-3 py-1.5 rounded-lg cursor-pointer"><i class="fa-solid fa-wand-magic-sparkles"></i> 食い違い${cmp.diff}か所をすべて画像の値に直す</button><span class="text-[11px] text-slate-500">赤いマスの「画像の値に」で1か所ずつ直すこともできます。画像の読み間違いならそのままにしてください。</span></div>`
            : (cmp && cmp.match ? `<div class="text-xs font-bold text-emerald-700"><i class="fa-solid fa-circle-check"></i> 読み取れたマスはすべてシフト表と一致しました。</div>` : '')}
          <details class="text-xs"><summary class="cursor-pointer font-bold text-slate-600">読み取った行と人の対応（違っていたら直せます）</summary><div class="mt-2 space-y-1.5">${mapRows || '<span class="text-slate-400">シフトらしい行が見つかりませんでした</span>'}</div></details>
          ${table}
          <details class="text-xs"><summary class="cursor-pointer text-slate-500">読み取った文字（そのまま）</summary><pre class="mt-1 p-2 bg-slate-50 rounded max-h-40 overflow-auto whitespace-pre-wrap">${esc(ocrState.rawText)}</pre></details>`;
    }
    el.innerHTML = `
      <div class="flex flex-wrap justify-between items-start gap-2">
        <h2 class="text-base font-bold text-slate-700 flex items-center gap-2"><i class="fa-solid fa-camera text-indigo-600"></i> 画像認識でシフトを二重チェック</h2>
        ${ocrImage ? `<button type="button" onclick="ocrClear()" class="text-xs px-2.5 py-1 rounded-md border border-slate-300 bg-white hover:bg-slate-50 cursor-pointer"><i class="fa-solid fa-xmark"></i> 画像を外す</button>` : ''}
      </div>
      <p class="text-xs text-slate-500">元のシフト表の<strong>スクリーンショット・写真</strong>を貼り付けると、画像の文字を読み取って、読み込んだシフトと1マスずつ照合します。<strong>日付の行（1, 2, 3…）と名前の列が写るように</strong>撮ってください。画像認識は読み間違えることもあるので、食い違いは目で確認してから直してください。</p>
      ${hasData ? '' : `<div class="text-xs font-bold text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">先に上でシフト表（Excel）を読み込んでください。照合する相手になります。</div>`}
      <div id="ocrDrop" tabindex="0" class="border-2 border-dashed border-slate-300 rounded-xl p-4 text-center text-xs text-slate-500 cursor-pointer hover:border-indigo-400 hover:bg-indigo-50/40 transition"
           onclick="document.getElementById('ocrFile').click()"
           ondragover="event.preventDefault(); this.classList.add('drag')" ondragleave="this.classList.remove('drag')"
           ondrop="event.preventDefault(); this.classList.remove('drag'); ocrSetImage(event.dataTransfer.files[0])">
        ${ocrImage ? `<img src="${ocrImage.src}" class="max-h-64 mx-auto rounded shadow" alt="貼り付けた画像">` :
          `<i class="fa-regular fa-image text-3xl text-slate-400 mb-1"></i><div class="font-bold text-slate-600">ここをクリックして画像を選ぶ ／ 画像をドラッグ ／ <kbd class="px-1 border rounded bg-white">Ctrl</kbd>+<kbd class="px-1 border rounded bg-white">V</kbd> で貼り付け</div>`}
      </div>
      <input type="file" id="ocrFile" accept="image/*" class="hidden" onchange="ocrSetImage(this.files[0]); this.value='';">
      <div class="flex flex-wrap items-center gap-3">
        <button type="button" onclick="runOcr()" ${ocrImage && hasData ? '' : 'disabled'} class="${ocrImage && hasData ? 'bg-indigo-600 hover:bg-indigo-700 cursor-pointer' : 'bg-slate-300 cursor-not-allowed'} text-white text-sm font-bold px-4 py-2 rounded-lg inline-flex items-center gap-2"><i class="fa-solid fa-magnifying-glass"></i> 読み取って照合する</button>
        <div id="ocrProg" class="flex-1 min-w-[12rem]"></div>
      </div>
      ${result}`;
}

// 画面0・画面1に二重チェックの欄を置く
(function () {
    const sec0 = document.getElementById('tab0');
    if (sec0) sec0.insertAdjacentHTML('beforeend', `<div id="ocrCard" class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 space-y-3"></div>`);
    const tab1Head = document.querySelector('#tab1 #loadStatus');
    if (tab1Head) tab1Head.parentElement.insertAdjacentHTML('afterbegin', `<button type="button" onclick="switchTab('tab0'); setTimeout(()=>document.getElementById('ocrCard').scrollIntoView({behavior:'smooth'}),80)" class="text-xs font-medium px-3 py-1 rounded-md border border-indigo-300 text-indigo-700 bg-white hover:bg-indigo-50 cursor-pointer"><i class="fa-solid fa-camera"></i> 画像で二重チェック</button>`);
    renderOcrCard();
})();
// どこで Ctrl+V しても、画面0か1を開いていれば画像を受け取る
document.addEventListener('paste', e => {
    const t0 = document.getElementById('tab0'), t1 = document.getElementById('tab1');
    const onTab = (t0 && !t0.classList.contains('hidden')) || (t1 && !t1.classList.contains('hidden'));
    if (!onTab) return;
    const item = [...((e.clipboardData && e.clipboardData.items) || [])].find(i => i.type && i.type.startsWith('image/'));
    if (!item) return;
    e.preventDefault();
    if (t1 && !t1.classList.contains('hidden')) switchTab('tab0');
    ocrSetImage(item.getAsFile());
    setTimeout(() => { const c = document.getElementById('ocrCard'); if (c) c.scrollIntoView({ behavior: 'smooth' }); }, 80);
});
// データが変わったら照合結果も更新
const _parseExcelP5 = parseExcelExactLayout;
parseExcelExactLayout = function () { const r = _parseExcelP5.apply(this, arguments); try { renderOcrCard(); } catch (e) { } return r; };
const _resetAllP5 = resetAll;
resetAll = function () { _resetAllP5(); if (!rawShiftData || !rawShiftData.length) { ocrImage = null; ocrState = null; renderOcrCard(); } };
</script>
<script>
// ===== 画像二重チェック：強化版（PNG/JPEG/WebP/BMP/GIF/PDF・複数枚・自動リトライ） =====
let ocrImages = [];          // { img(canvas or Image), name, url }
let ocrStartDay = 1;         // 日付の行が読めないときの「左端の列の日付」
let ocrNote = "";            // 画面に出すお知らせ・エラー

const OCR_IMG_EXT = /\.(png|jpe?g|jfif|gif|bmp|webp|tiff?|avif|heic|heif|ico)$/i;

function loadScriptOnce(src, globalName) {
    if (globalName && window[globalName]) return Promise.resolve();
    return new Promise((ok, ng) => {
        const s = document.createElement('script');
        s.src = src; s.onload = () => ok(); s.onerror = () => ng(new Error("部品を読み込めませんでした（インターネット接続を確認してください）"));
        document.head.appendChild(s);
    });
}

function fileToDataURL(file) {
    return new Promise((ok, ng) => { const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.onerror = () => ng(fr.error); fr.readAsDataURL(file); });
}
function loadImageFrom(src) {
    return new Promise((ok, ng) => { const im = new Image(); im.onload = () => ok(im); im.onerror = () => ng(new Error("decode")); im.src = src; });
}
// PDFは各ページを画像にする
async function pdfToCanvases(file) {
    await loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js", "pdfjsLib");
    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const out = [];
    for (let p = 1; p <= Math.min(pdf.numPages, 6); p++) {
        const page = await pdf.getPage(p);
        const vp = page.getViewport({ scale: 2.5 });
        const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
        await page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
        out.push(cv);
    }
    return out;
}

async function ocrAddFiles(files) {
    files = [...(files || [])].filter(Boolean);
    if (!files.length) return;
    let added = 0;
    const bad = [];
    for (const f of files) {
        const name = f.name || "貼り付け画像";
        try {
            if (/pdf/i.test(f.type) || /\.pdf$/i.test(name)) {
                const cvs = await pdfToCanvases(f);
                cvs.forEach((cv, i) => { ocrImages.push({ img: cv, name: `${name}（${i + 1}ページ）`, url: cv.toDataURL('image/png') }); added++; });
                continue;
            }
            if (f.type && !f.type.startsWith("image/") && !OCR_IMG_EXT.test(name)) { bad.push(`${name}（画像ではありません）`); continue; }
            const url = await fileToDataURL(f);
            let img;
            try { img = await loadImageFrom(url); }
            catch (e) {
                // 拡張子と中身が違う・種類が空などでも、もう一度ビットマップとして読んでみる
                if (window.createImageBitmap) {
                    const bmp = await createImageBitmap(f);
                    const cv = document.createElement('canvas'); cv.width = bmp.width; cv.height = bmp.height; cv.getContext('2d').drawImage(bmp, 0, 0);
                    img = cv;
                } else throw e;
            }
            ocrImages.push({ img, name, url: img instanceof HTMLCanvasElement ? img.toDataURL('image/png') : url });
            added++;
        } catch (e) {
            console.error(e);
            bad.push(/heic|heif/i.test(name) ? `${name}（HEIC形式はこのブラウザで開けません。PNGかJPEGに変換してください）` : `${name}（読み込めない形式です。PNG・JPEGで保存し直してください）`);
        }
    }
    ocrState = null;
    ocrNote = bad.length ? "⚠ 読み込めなかったファイル：" + bad.join("、") : "";
    renderOcrCard();
    if (added) toast(`画像を${added}枚読み込みました。「読み取って照合する」を押してください`, true);
}
ocrSetImage = function (file) { ocrAddFiles([file]); };
function ocrRemoveImage(i) { ocrImages.splice(i, 1); ocrState = null; renderOcrCard(); }
ocrClear = function () { ocrImages = []; ocrState = null; ocrNote = ""; renderOcrCard(); };

// ---- 画像の下ごしらえ ----
function imgSize(img) { return { w: img.naturalWidth || img.width, h: img.naturalHeight || img.height }; }
function ocrCanvas(img, mode) {
    const { w, h } = imgSize(img);
    let scale = Math.max(1, Math.min(4, 3200 / w));
    if (w * h * scale * scale > 24e6) scale = Math.sqrt(24e6 / (w * h));
    const cv = document.createElement('canvas');
    cv.width = Math.max(1, Math.round(w * scale)); cv.height = Math.max(1, Math.round(h * scale));
    const ctx = cv.getContext('2d');
    ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height);      // 透過PNGの背景を白に
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(img, 0, 0, cv.width, cv.height);
    const d = ctx.getImageData(0, 0, cv.width, cv.height), p = d.data;
    const gray = new Uint8ClampedArray(p.length / 4);
    let lo = 255, hi = 0;
    for (let i = 0, j = 0; i < p.length; i += 4, j++) { const g = 0.299 * p[i] + 0.587 * p[i + 1] + 0.114 * p[i + 2]; gray[j] = g; }
    // 明るさの幅を広げる（コントラスト強調）
    const hist = new Array(256).fill(0); gray.forEach(g => hist[g]++);
    let acc = 0; const n = gray.length;
    for (let i = 0; i < 256; i++) { acc += hist[i]; if (acc > n * 0.01) { lo = i; break; } }
    acc = 0; for (let i = 255; i >= 0; i--) { acc += hist[i]; if (acc > n * 0.01) { hi = i; break; } }
    const span = Math.max(1, hi - lo);
    let th = 128;
    if (mode === 'bin') {   // 大津の方法で白黒に
        const h2 = new Array(256).fill(0);
        gray.forEach(g => h2[Math.max(0, Math.min(255, Math.round((g - lo) * 255 / span)))]++);
        let sum = 0; for (let i = 0; i < 256; i++) sum += i * h2[i];
        let sumB = 0, wB = 0, best = 0;
        for (let i = 0; i < 256; i++) { wB += h2[i]; if (!wB) continue; const wF = n - wB; if (!wF) break; sumB += i * h2[i]; const mB = sumB / wB, mF = (sum - sumB) / wF; const v = wB * wF * (mB - mF) * (mB - mF); if (v > best) { best = v; th = i; } }
    }
    for (let i = 0, j = 0; i < p.length; i += 4, j++) {
        let g = Math.max(0, Math.min(255, (gray[j] - lo) * 255 / span));
        if (mode === 'bin') g = g < th ? 0 : 255;
        p[i] = p[i + 1] = p[i + 2] = g;
    }
    ctx.putImageData(d, 0, 0);
    return cv;
}

// ---- 読み取り結果を表に並べ直す（日付の行が無くても列を推定） ----
ocrAnalyze = function (words, startDay) {
    if (!words.length) return { error: "文字を読み取れませんでした。", rows: [], wordCount: 0 };
    const hs = words.map(w => w.h).sort((a, b) => a - b), mh = hs[Math.floor(hs.length / 2)] || 10;
    const rows = [];
    words.slice().sort((a, b) => a.cy - b.cy).forEach(w => {
        const row = rows.find(r => Math.abs(r.cy - w.cy) < mh * 0.6);
        if (row) { row.w.push(w); row.cy = row.w.reduce((t, x) => t + x.cy, 0) / row.w.length; } else rows.push({ cy: w.cy, w: [w] });
    });
    rows.forEach(r => r.w.sort((a, b) => a.cx - b.cx));

    const dayOf = t => {
        const s = toHalfWidth(t).replace(/[\s日]/g, "").replace(/[\(（][^)）]*[\)）]?$/, "").replace(/[月火水木金土]$/, "");
        let m = s.match(/^\d{1,2}[\/\.](\d{1,2})$/); if (m && +m[1] >= 1 && +m[1] <= 31) return +m[1];
        m = s.match(/^(\d{1,2})$/); if (m && +m[1] >= 1 && +m[1] <= 31) return +m[1];
        return null;
    };
    let header = null;
    rows.forEach(r => {
        const ds = r.w.map(w => ({ d: dayOf(w.t), x: w.cx })).filter(o => o.d !== null);
        let inc = 0; for (let i = 1; i < ds.length; i++) if (ds[i].d > ds[i - 1].d) inc++;
        if (ds.length >= 3 && inc >= (ds.length - 1) * 0.7 && (!header || ds.length > header.ds.length)) header = { row: r, ds };
    });

    let dayX = {}, headerFound = !!header, topY = -Infinity;
    if (header) {
        header.ds.forEach(o => { if (dayX[o.d] === undefined) dayX[o.d] = o.x; });
        topY = header.row.cy;
        // 読み落とした日付の列を、前後の列から補う
        const ks = Object.keys(dayX).map(Number).sort((a, b) => a - b);
        for (let i = 1; i < ks.length; i++) {
            const a = ks[i - 1], b = ks[i];
            for (let d = a + 1; d < b; d++) dayX[d] = dayX[a] + (dayX[b] - dayX[a]) * (d - a) / (b - a);
        }
    } else {
        // 日付の行が無い：シフトらしい文字のX位置から列を推定し、左から startDay, startDay+1 … とみなす
        const xs = [];
        rows.forEach(r => { const sw = r.w.filter(w => normShift(w.t)); if (sw.length >= 3) sw.forEach(w => xs.push(w.cx)); });
        xs.sort((a, b) => a - b);
        const cols = [];
        xs.forEach(x => { const c = cols[cols.length - 1]; if (c && x - c.max < mh * 1.3) { c.max = x; c.sum += x; c.n++; } else cols.push({ min: x, max: x, sum: x, n: 1 }); });
        const good = cols.filter(c => c.n >= 2);
        good.forEach((c, i) => { dayX[(startDay || 1) + i] = c.sum / c.n; });
        if (!good.length) return { error: "日付の行も、シフトの列も見つかりませんでした。", rows: [], wordCount: words.length };
    }
    const xs = Object.values(dayX).sort((a, b) => a - b);
    const gaps = []; for (let i = 1; i < xs.length; i++) gaps.push(xs[i] - xs[i - 1]);
    gaps.sort((a, b) => a - b);
    const colW = gaps.length ? gaps[Math.floor(gaps.length / 2)] : mh * 3;
    const entries = Object.entries(dayX).map(([d, x]) => ({ d: +d, x }));
    const leftEdge = Math.min(...xs) - colW * 0.6;

    const out = [];
    rows.forEach(r => {
        if (header && (r === header.row || r.cy <= topY)) return;
        const nameText = r.w.filter(w => w.cx < leftEdge).map(w => w.t).join("");
        const cells = {};
        r.w.filter(w => w.cx >= leftEdge).forEach(w => {
            let best = null;
            entries.forEach(e => { const dist = Math.abs(e.x - w.cx); if (dist < colW * 0.6 && (!best || dist < best.dist)) best = { d: e.d, dist }; });
            if (best) (cells[best.d] = cells[best.d] || []).push(w);
        });
        const cellText = {};
        Object.keys(cells).forEach(d => { cellText[d] = cells[d].sort((a, b) => a.cx - b.cx).map(w => w.t).join(" "); });
        if (Object.values(cellText).filter(t => normShift(t)).length >= 2) out.push({ nameText, cells: cellText, staff: "" });
    });
    const used = new Set();
    out.forEach(row => {
        const nt = row.nameText.replace(/\s/g, "");
        let hit = staffList.find(s => !used.has(s.name) && [s.name, ...(s.aliases || [])].some(k => k && nt.includes(k)));
        if (!hit) hit = staffList.find(s => !used.has(s.name) && [s.name, ...(s.aliases || [])].some(k => k && k.length >= 2 && [...new Set(k)].filter(ch => nt.includes(ch)).length >= 2));
        if (!hit && nt) hit = staffList.find(s => !used.has(s.name) && [s.name, ...(s.aliases || [])].some(k => k && [...k].some(ch => /[一-鿿]/.test(ch) && nt.includes(ch))));
        if (hit) { row.staff = hit.name; used.add(hit.name); }
    });
    return { rows: out, headerFound, wordCount: words.length, days: Object.keys(dayX).map(Number).sort((a, b) => a - b) };
};

function ocrQuality(res) {
    if (!res || !res.rows) return -1;
    let n = 0; res.rows.forEach(r => Object.values(r.cells).forEach(t => { if (normShift(t)) n++; }));
    return n + (res.headerFound ? 50 : 0) + res.rows.filter(r => r.staff).length * 10;
}

runOcr = async function () {
    if (!ocrImages.length) { toast("先に画像を貼り付けてください"); return; }
    if (!rawShiftData || !rawShiftData.length) { toast("先にシフト表（Excel）を読み込んでください"); return; }
    const prog = document.getElementById('ocrProg');
    const setProg = (t, v) => { if (prog) prog.innerHTML = `<div class="text-xs text-slate-600 mb-1">${esc(t)}</div><div class="h-2 rounded bg-slate-200 overflow-hidden"><div class="h-full bg-indigo-500 transition-all" style="width:${Math.round(Math.max(0, Math.min(1, v || 0)) * 100)}%"></div></div>`; };
    let curLabel = "";
    try {
        setProg("画像認識の準備中…（初回は日本語データの読み込みに少し時間がかかります）", 0.03);
        await loadScriptOnce("https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js", "Tesseract");
        if (!ocrWorker) {
            ocrWorker = await Tesseract.createWorker('jpn', 1, {
                logger: m => { if (m.status === 'recognizing text') setProg(`${curLabel} 文字を読み取り中…`, m.progress); else if (m.status) setProg("準備中…（" + m.status + "）", m.progress); }
            });
        }
        const all = { rows: [], notes: [] };
        for (let i = 0; i < ocrImages.length; i++) {
            const im = ocrImages[i];
            curLabel = `画像${i + 1}/${ocrImages.length}`;
            // 1回目：コントラスト強調・まばらな文字モード
            await ocrWorker.setParameters({ tessedit_pageseg_mode: '11', preserve_interword_spaces: '1' });
            const r1 = await ocrWorker.recognize(ocrCanvas(im.img, 'gray'));
            let best = ocrAnalyze(ocrWordsOf(r1.data), ocrStartDay); best.rawText = r1.data.text || "";
            // 読み取りが弱ければ2回目：白黒化・表モードで読み直し、良い方を採用
            if (!best.headerFound || ocrQuality(best) < 60) {
                curLabel = `画像${i + 1}/${ocrImages.length}（読み直し）`;
                await ocrWorker.setParameters({ tessedit_pageseg_mode: '6', preserve_interword_spaces: '1' });
                const r2 = await ocrWorker.recognize(ocrCanvas(im.img, 'bin'));
                const b2 = ocrAnalyze(ocrWordsOf(r2.data), ocrStartDay); b2.rawText = r2.data.text || "";
                if (ocrQuality(b2) > ocrQuality(best)) best = b2;
            }
            (best.rows || []).forEach(r => { r.img = i + 1; all.rows.push(r); });
            all.notes.push(best.error ? `画像${i + 1}：${best.error}` : `画像${i + 1}：文字${best.wordCount}個・シフト行${best.rows.length}行${best.headerFound ? '' : '（日付の行が読めなかったため、左端の列を' + ocrStartDay + '日として推定）'}`);
            all.rawText = (all.rawText || "") + `【画像${i + 1}】\n` + (best.rawText || "") + "\n";
        }
        // 同じ人が複数の画像に出てきたら、2つ目以降は自動対応を外して重複を避ける
        const seen = new Set();
        all.rows.forEach(r => { if (r.staff) { if (seen.has(r.staff + "|" + r.img)) r.staff = ""; else seen.add(r.staff + "|" + r.img); } });
        ocrState = { rows: all.rows, rawText: all.rawText, notes: all.notes, error: all.rows.length ? "" : "シフトの行を読み取れませんでした。" };
        if (prog) prog.innerHTML = "";
        renderOcrCard();
    } catch (err) {
        console.error(err);
        if (prog) prog.innerHTML = `<div class="text-xs font-bold text-red-600">${esc(err.message || "画像認識でエラーが発生しました")}</div>`;
    }
};

ocrCompare = function () {
    if (!ocrState || !ocrState.rows) return null;
    const dataOf = (n, d) => { const it = (rawShiftData || []).find(x => x.name === n && parseInt(String(x.date).split("/").pop(), 10) === d); return it ? { v: it.shiftType, date: it.date } : null; };
    const res = { match: 0, diff: 0, unread: 0, rows: [], days: [] };
    const daySet = new Set();
    ocrState.rows.forEach(row => { if (row.staff) Object.keys(row.cells).forEach(d => daySet.add(+d)); });
    res.days = [...daySet].sort((a, b) => a - b);
    ocrState.rows.forEach((row, ri) => {
        if (!row.staff) return;
        const cells = res.days.map(d => {
            const dv = dataOf(row.staff, d);
            const raw = row.cells[d];
            if (!dv || raw === undefined) return { d, state: "none" };
            const k = normShift(raw);
            if (!k) { res.unread++; return { d, date: dv.date, data: dv.v, raw, state: "unread" }; }
            if (k === normData(dv.v)) { res.match++; return { d, date: dv.date, data: dv.v, raw, state: "match" }; }
            res.diff++;
            return { d, date: dv.date, data: dv.v, raw, img: normLabel(k, raw), state: "diff" };
        });
        res.rows.push({ ri, staff: row.staff, img: row.img, cells });
    });
    return res;
};

renderOcrCard = function () {
    const el = document.getElementById('ocrCard'); if (!el) return;
    const hasData = rawShiftData && rawShiftData.length;
    const cmp = ocrCompare();
    let result = "";
    if (ocrState) {
        const notes = (ocrState.notes || []).map(n => `<li>${esc(n)}</li>`).join('');
        const mapRows = ocrState.rows.map((r, i) => `<div class="flex flex-wrap items-center gap-2 text-xs"><span class="text-slate-400 whitespace-nowrap">画像${r.img}・行${i + 1}</span><span class="px-1.5 py-0.5 rounded bg-slate-100 font-mono max-w-[14rem] truncate" title="${esc(r.nameText)}">${esc(r.nameText || "（名前読めず）")}</span><i class="fa-solid fa-arrow-right text-slate-300"></i>
            <select onchange="ocrSetStaff(${i}, this.value)" class="border border-slate-300 rounded px-1.5 py-0.5 text-xs ${r.staff ? 'font-bold text-indigo-700' : 'text-slate-400'}"><option value="">（照合しない）</option>${staffList.map(s => `<option value="${esc(s.name)}" ${s.name === r.staff ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select>
            <span class="text-slate-400 truncate max-w-[18rem]" title="${esc(Object.entries(r.cells).map(([d, t]) => d + '日:' + t).join(' '))}">${Object.entries(r.cells).slice(0, 6).map(([d, t]) => d + '日:' + esc(t)).join(' ')}…</span></div>`).join('');
        const unmapped = ocrState.rows.filter(r => !r.staff).length;
        const table = cmp && cmp.rows.length ? `
          <div class="overflow-auto border border-slate-200 rounded-lg max-h-[50vh]">
            <table class="compact-table text-xs text-center border-separate border-spacing-0">
              <thead class="sticky top-0 bg-slate-100 z-10"><tr><th class="px-2 py-1 sticky left-0 bg-slate-100">氏名</th>${cmp.days.map(d => `<th class="px-1 py-1 min-w-[3.2rem]">${d}日</th>`).join('')}</tr></thead>
              <tbody>${cmp.rows.map(r => `<tr><td class="px-2 py-1 font-bold sticky left-0 bg-white whitespace-nowrap">${esc(r.staff)}${ocrImages.length > 1 ? `<span class="text-[10px] text-slate-400 font-normal ml-1">画像${r.img}</span>` : ''}</td>${r.cells.map(x => {
                if (x.state === "none") return `<td class="px-1 py-1 text-slate-300">-</td>`;
                if (x.state === "match") return `<td class="px-1 py-1 bg-emerald-50 text-emerald-800 font-bold border border-white" title="画像：${esc(x.raw)}">${esc(x.data)} <i class="fa-solid fa-check text-[9px]"></i></td>`;
                if (x.state === "unread") return `<td class="px-1 py-1 bg-slate-50 text-slate-400 border border-white" title="画像の文字「${esc(x.raw)}」をシフトとして判定できませんでした">${esc(x.data)}<div class="text-[9px]">読取不可</div></td>`;
                return `<td class="px-1 py-1 bg-red-100 border-2 border-red-400 text-red-800" title="データ：${esc(x.data)} ／ 画像：${esc(x.raw)}"><div class="font-bold line-through decoration-red-400">${esc(x.data)}</div><div class="text-[10px]">画像:${esc(x.img)}</div>
                    <button type="button" onclick="ocrApply('${esc(r.staff)}','${esc(x.date)}','${esc(x.img)}')" class="mt-0.5 px-1 rounded bg-red-500 hover:bg-red-600 text-white text-[10px] cursor-pointer">画像の値に</button></td>`;
              }).join('')}</tr>`).join('')}</tbody>
            </table>
          </div>` : "";
        result = `
          ${notes ? `<ul class="text-[11px] text-slate-500 list-disc list-inside">${notes}</ul>` : ''}
          ${ocrState.error ? `<div class="p-3 rounded-lg bg-red-50 border border-red-200 text-xs text-red-700 space-y-1"><div class="font-bold"><i class="fa-solid fa-circle-exclamation"></i> ${esc(ocrState.error)}</div>
              <div>うまく読めないときは：①表の部分だけを大きめに切り取る ②日付の行（1, 2, 3…）と名前の列を入れる ③ぼやけた写真は避けてスクリーンショットを使う ④何枚かに分けて貼る（複数枚OK）</div></div>` : ''}
          ${cmp && cmp.rows.length ? `
          <div class="grid grid-cols-3 gap-2 text-center">
            <div class="p-2 rounded-lg bg-emerald-50 border border-emerald-200"><div class="text-[11px] text-emerald-700">一致</div><div class="text-lg font-bold text-emerald-700">${cmp.match}</div></div>
            <div class="p-2 rounded-lg bg-red-50 border border-red-200"><div class="text-[11px] text-red-700">食い違い</div><div class="text-lg font-bold text-red-700">${cmp.diff}</div></div>
            <div class="p-2 rounded-lg bg-slate-50 border border-slate-200"><div class="text-[11px] text-slate-500">読み取れず</div><div class="text-lg font-bold text-slate-500">${cmp.unread}</div></div>
          </div>
          ${cmp.diff ? `<div class="flex flex-wrap items-center gap-2"><button type="button" onclick="ocrApplyAll()" class="bg-red-500 hover:bg-red-600 text-white text-xs font-bold px-3 py-1.5 rounded-lg cursor-pointer"><i class="fa-solid fa-wand-magic-sparkles"></i> 食い違い${cmp.diff}か所をすべて画像の値に直す</button><span class="text-[11px] text-slate-500">赤いマスの「画像の値に」で1か所ずつも直せます。画像の読み間違いならそのままに。</span></div>`
            : (cmp.match ? `<div class="text-xs font-bold text-emerald-700"><i class="fa-solid fa-circle-check"></i> 読み取れたマスはすべてシフト表と一致しました。</div>` : '')}` : ''}
          ${ocrState.rows.length ? `<details ${unmapped || !(cmp && cmp.rows.length) ? 'open' : ''} class="text-xs"><summary class="cursor-pointer font-bold text-slate-600">読み取った行と人の対応${unmapped ? `（<span class="text-amber-700">名前が読めなかった行が${unmapped}行あります。誰の行か選んでください</span>）` : '（違っていたら直せます）'}</summary><div class="mt-2 space-y-1.5">${mapRows}</div></details>` : ''}
          ${table}
          <details class="text-xs"><summary class="cursor-pointer text-slate-500">読み取った文字（そのまま）</summary><pre class="mt-1 p-2 bg-slate-50 rounded max-h-40 overflow-auto whitespace-pre-wrap">${esc(ocrState.rawText || "")}</pre></details>`;
    }
    const thumbs = ocrImages.map((im, i) => `
        <div class="relative group">
          <img src="${im.url}" class="h-28 max-w-[16rem] object-contain rounded border border-slate-200 bg-white shadow-sm" alt="${esc(im.name)}">
          <div class="text-[10px] text-slate-500 truncate max-w-[16rem]">画像${i + 1}：${esc(im.name)}</div>
          <button type="button" onclick="event.stopPropagation(); ocrRemoveImage(${i})" title="外す" class="absolute -top-2 -right-2 w-6 h-6 rounded-full bg-white border border-slate-300 text-slate-500 hover:text-red-600 shadow cursor-pointer"><i class="fa-solid fa-xmark"></i></button>
        </div>`).join('');
    el.innerHTML = `
      <div class="flex flex-wrap justify-between items-start gap-2">
        <h2 class="text-base font-bold text-slate-700 flex items-center gap-2"><i class="fa-solid fa-camera text-indigo-600"></i> 画像認識でシフトを二重チェック</h2>
        ${ocrImages.length ? `<button type="button" onclick="ocrClear()" class="text-xs px-2.5 py-1 rounded-md border border-slate-300 bg-white hover:bg-slate-50 cursor-pointer"><i class="fa-solid fa-xmark"></i> 画像をすべて外す</button>` : ''}
      </div>
      <p class="text-xs text-slate-500">元のシフト表の<strong>スクリーンショット・写真・PDF</strong>を入れると、文字を読み取って、読み込んだシフトと1マスずつ照合します。PNG・JPEG・WebP・BMP・GIF・PDFに対応、<strong>何枚でも</strong>まとめて入れられます。<strong>日付の行（1, 2, 3…）と名前の列</strong>が写っていると精度が上がります。</p>
      ${hasData ? '' : `<div class="text-xs font-bold text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">先に上でシフト表（Excel）を読み込んでください。照合する相手になります。</div>`}
      ${ocrNote ? `<div class="text-xs font-bold text-amber-800 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">${esc(ocrNote)}</div>` : ''}
      <div id="ocrDrop" tabindex="0" class="border-2 border-dashed border-slate-300 rounded-xl p-4 text-center text-xs text-slate-500 cursor-pointer hover:border-indigo-400 hover:bg-indigo-50/40 transition"
           onclick="document.getElementById('ocrFile').click()"
           ondragover="event.preventDefault(); this.classList.add('drag')" ondragleave="this.classList.remove('drag')"
           ondrop="ocrHandleDrop(event)">
        ${ocrImages.length ? `<div class="flex flex-wrap gap-4 justify-center">${thumbs}</div><div class="mt-3 text-indigo-600 font-bold"><i class="fa-solid fa-plus"></i> クリック・ドラッグ・Ctrl+Vで画像を追加</div>` :
          `<i class="fa-regular fa-image text-3xl text-slate-400 mb-1"></i><div class="font-bold text-slate-600">ここをクリックして画像を選ぶ ／ 画像をドラッグ ／ <kbd class="px-1 border rounded bg-white">Ctrl</kbd>+<kbd class="px-1 border rounded bg-white">V</kbd> で貼り付け</div><div class="mt-1">PNG・JPEG・WebP・BMP・GIF・PDF（複数OK）</div>`}
      </div>
      <input type="file" id="ocrFile" accept="image/*,.png,.jpg,.jpeg,.jfif,.gif,.bmp,.webp,.tif,.tiff,.avif,.heic,.heif,.pdf,application/pdf" multiple class="hidden" onchange="ocrAddFiles(this.files); this.value='';">
      <div class="flex flex-wrap items-center gap-3">
        <button type="button" onclick="runOcr()" ${ocrImages.length && hasData ? '' : 'disabled'} class="${ocrImages.length && hasData ? 'bg-indigo-600 hover:bg-indigo-700 cursor-pointer' : 'bg-slate-300 cursor-not-allowed'} text-white text-sm font-bold px-4 py-2 rounded-lg inline-flex items-center gap-2"><i class="fa-solid fa-magnifying-glass"></i> 読み取って照合する</button>
        <label class="text-xs text-slate-500 flex items-center gap-1" title="画像に日付の行が写っていない・読めないときに使います">日付の行が無い場合、左端の列は<input type="number" min="1" max="31" value="${ocrStartDay}" onchange="ocrStartDay=Math.max(1,Math.min(31,parseInt(this.value,10)||1))" class="w-12 border border-slate-300 rounded px-1 py-0.5">日</label>
        <div id="ocrProg" class="flex-1 min-w-[12rem]"></div>
      </div>
      ${result}`;
};

function ocrHandleDrop(e) {
    e.preventDefault();
    const dz = document.getElementById('ocrDrop'); if (dz) dz.classList.remove('drag');
    const dt = e.dataTransfer;
    if (dt.files && dt.files.length) { ocrAddFiles(dt.files); return; }
    // ブラウザ上の画像をドラッグした場合
    const url = dt.getData('text/uri-list') || dt.getData('text/plain');
    if (url && /^(https?:|data:image)/.test(url)) {
        const im = new Image(); im.crossOrigin = "anonymous";
        im.onload = () => {
            try { const cv = document.createElement('canvas'); cv.width = im.naturalWidth; cv.height = im.naturalHeight; cv.getContext('2d').drawImage(im, 0, 0); cv.toDataURL(); ocrImages.push({ img: cv, name: "ドラッグした画像", url: cv.toDataURL('image/png') }); ocrState = null; ocrNote = ""; renderOcrCard(); }
            catch (err) { ocrNote = "⚠ この画像は直接読み込めません。一度パソコンに保存してから入れるか、スクリーンショットを貼り付けてください。"; renderOcrCard(); }
        };
        im.onerror = () => { ocrNote = "⚠ この画像は直接読み込めません。一度パソコンに保存してから入れるか、スクリーンショットを貼り付けてください。"; renderOcrCard(); };
        im.src = url;
    } else { ocrNote = "⚠ 画像が見つかりませんでした。画像ファイルをドラッグしてください。"; renderOcrCard(); }
}

// どの画面でも Ctrl+V で画像を受け取る（複数OK）
document.addEventListener('paste', e => {
    const items = [...((e.clipboardData && e.clipboardData.items) || [])].filter(i => i.kind === 'file');
    const files = items.map(i => i.getAsFile()).filter(f => f && (!f.type || f.type.startsWith('image/') || /pdf/i.test(f.type) || OCR_IMG_EXT.test(f.name || "")));
    if (!files.length) return;
    e.preventDefault(); e.stopImmediatePropagation();
    switchTab('tab0');
    ocrAddFiles(files);
    setTimeout(() => { const c = document.getElementById('ocrCard'); if (c) c.scrollIntoView({ behavior: 'smooth' }); }, 80);
}, true);
renderOcrCard();
</script>
<script>
const _resetAllP6 = resetAll;
resetAll = function () { _resetAllP6(); if (!rawShiftData || !rawShiftData.length) { ocrImages = []; ocrState = null; ocrNote = ""; renderOcrCard(); } };
</script>
<script>
// ===== 画像認識：Worker が使えない環境（claude.ai のプレビュー等）でも動くようにする =====
// 1) まず通常の Tesseract.js（Worker）を試す
// 2) Worker が作れない場合は、文字認識エンジン本体（WebAssembly）をページ上で直接動かす
let ocrEngine = null;      // { kind: 'worker'|'main', recognize(canvas, psm, onProg) → {words, text} }
const TESS_CORE_URL = "https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.1/tesseract-core-lstm.wasm.js";
const TESS_LANG_URL = "https://cdn.jsdelivr.net/npm/@tesseract.js-data/jpn/4.0.0_best_int/jpn.traineddata.gz";

async function gunzipIfNeeded(buf) {
    const u8 = new Uint8Array(buf);
    if (!(u8[0] === 0x1f && u8[1] === 0x8b)) return u8;
    if (!window.DecompressionStream) throw new Error("このブラウザは日本語データの展開に対応していません（ChromeかEdgeの最新版を使ってください）");
    const ds = new Blob([u8]).stream().pipeThrough(new DecompressionStream('gzip'));
    return new Uint8Array(await new Response(ds).arrayBuffer());
}
function canvasToPng(cv) {
    return new Promise((ok, ng) => cv.toBlob(b => b ? b.arrayBuffer().then(a => ok(new Uint8Array(a))) : ng(new Error("画像の変換に失敗しました")), 'image/png'));
}
// TSV（1行=1単語、位置つき）→ ocrWordsOf と同じ形
function tsvToWords(tsv) {
    const words = [];
    String(tsv || "").split("\n").forEach(line => {
        const c = line.split("\t");
        if (c.length < 12 || c[0] !== "5") return;
        const t = c.slice(11).join("\t").trim(); if (!t) return;
        const x = +c[6], y = +c[7], w = +c[8], h = +c[9];
        words.push({ text: t, bbox: { x0: x, y0: y, x1: x + w, y1: y + h }, confidence: +c[10] });
    });
    return words;
}
const nextFrame = () => new Promise(r => setTimeout(r, 30));

async function makeWorkerEngine(setProg) {
    await loadScriptOnce("https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js", "Tesseract");
    let progCb = null;
    const w = await Tesseract.createWorker('jpn', 1, { logger: m => { if (progCb && m.status === 'recognizing text') progCb(m.progress); else if (m.status) setProg("準備中…（" + m.status + "）", m.progress); } });
    return {
        kind: 'worker',
        async recognize(cv, psm, onProg) {
            progCb = onProg;
            await w.setParameters({ tessedit_pageseg_mode: String(psm), preserve_interword_spaces: '1' });
            const { data } = await w.recognize(cv);
            return { words: data.words || [], blocks: data.blocks, text: data.text || "" };
        }
    };
}

async function makeMainEngine(setProg) {
    setProg("文字認識エンジンを読み込み中…（このページ上で直接動かします）", 0.1);
    await loadScriptOnce(TESS_CORE_URL, "TesseractCore");
    if (!window.TesseractCore) throw new Error("文字認識エンジンを読み込めませんでした");
    const Module = await TesseractCore();
    setProg("日本語データを読み込み中…（初回は数十秒かかることがあります）", 0.3);
    const res = await fetch(TESS_LANG_URL);
    if (!res.ok) throw new Error("日本語データを読み込めませんでした（" + res.status + "）");
    const data = await gunzipIfNeeded(await res.arrayBuffer());
    Module.FS.writeFile("./jpn.traineddata", data);
    const api = new Module.TessBaseAPI();
    let st;
    try { st = api.Init(null, "jpn", 1); } catch (e) { st = api.Init(null, "jpn"); }
    if (st === -1) throw new Error("文字認識エンジンの初期化に失敗しました");
    setProg("準備完了", 1);
    return {
        kind: 'main',
        async recognize(cv, psm, onProg) {
            onProg && onProg(0.2);
            await nextFrame();                                   // 進み具合の表示を先に出す
            const png = await canvasToPng(cv);
            Module.FS.writeFile("/input", png);
            api.SetVariable("tessedit_pageseg_mode", String(psm));
            api.SetVariable("preserve_interword_spaces", "1");
            if (api.SetPageSegMode) { try { api.SetPageSegMode(Number(psm)); } catch (e) { } }
            let r;
            try { r = api.SetImageFile(1, 0); } catch (e) { r = api.SetImageFile(); }
            if (r === 1) throw new Error("画像を読み込めませんでした");
            onProg && onProg(0.5);
            await nextFrame();
            api.Recognize(null);
            const call = (fn, ...args) => { try { return api[fn](...args); } catch (e) { return api[fn](); } };
            const text = call("GetUTF8Text");
            let tsv = "";
            try { tsv = api.GetTSVText(0); } catch (e) { tsv = api.GetTSVText(); }
            try { api.Clear(); Module.FS.unlink("/input"); } catch (e) { }
            onProg && onProg(1);
            return { words: tsvToWords(tsv), text: text || "" };
        }
    };
}

async function getOcrEngine(setProg) {
    if (ocrEngine) return ocrEngine;
    let firstErr = null;
    try {
        if (typeof Worker === "undefined") throw new Error("no worker");
        ocrEngine = await makeWorkerEngine(setProg);
        return ocrEngine;
    } catch (e) {
        firstErr = e;
        console.warn("Worker版が使えないため、ページ上で直接動かします:", e);
    }
    try {
        ocrEngine = await makeMainEngine(setProg);
        return ocrEngine;
    } catch (e2) {
        console.error(e2);
        const msg = String(e2 && e2.message || e2);
        throw new Error("この画面では画像認識を動かせませんでした（" + msg + "）。HTMLファイルをダウンロードして、パソコンのChrome・Edgeで直接開くと使えます。" + (firstErr ? "" : ""));
    }
}

runOcr = async function () {
    if (!ocrImages.length) { toast("先に画像を貼り付けてください"); return; }
    if (!rawShiftData || !rawShiftData.length) { toast("先にシフト表（Excel）を読み込んでください"); return; }
    const prog = document.getElementById('ocrProg');
    const setProg = (t, v) => { if (prog) prog.innerHTML = `<div class="text-xs text-slate-600 mb-1">${esc(t)}</div><div class="h-2 rounded bg-slate-200 overflow-hidden"><div class="h-full bg-indigo-500 transition-all" style="width:${Math.round(Math.max(0, Math.min(1, v || 0)) * 100)}%"></div></div>`; };
    const btn = document.querySelector('#ocrCard button[onclick="runOcr()"]');
    if (btn) { btn.disabled = true; btn.classList.add('opacity-60'); }
    try {
        setProg("画像認識の準備中…（初回は日本語データの読み込みに少し時間がかかります）", 0.03);
        await nextFrame();
        const eng = await getOcrEngine(setProg);
        const slow = eng.kind === 'main' ? "（このページ上で処理中。画面が数十秒止まって見えることがありますが、そのままお待ちください）" : "";
        const all = { rows: [], notes: [], rawText: "" };
        for (let i = 0; i < ocrImages.length; i++) {
            const im = ocrImages[i];
            const lab = `画像${i + 1}/${ocrImages.length}`;
            setProg(`${lab} 文字を読み取り中…${slow}`, 0.05); await nextFrame();
            const r1 = await eng.recognize(ocrCanvas(im.img, 'gray'), 11, p => setProg(`${lab} 文字を読み取り中…${slow}`, p));
            let best = ocrAnalyze(ocrWordsOf(r1), ocrStartDay); best.rawText = r1.text;
            if (!best.headerFound || ocrQuality(best) < 60) {
                setProg(`${lab}（読み直し）文字を読み取り中…${slow}`, 0.05); await nextFrame();
                const r2 = await eng.recognize(ocrCanvas(im.img, 'bin'), 6, p => setProg(`${lab}（読み直し）文字を読み取り中…${slow}`, p));
                const b2 = ocrAnalyze(ocrWordsOf(r2), ocrStartDay); b2.rawText = r2.text;
                if (ocrQuality(b2) > ocrQuality(best)) best = b2;
            }
            (best.rows || []).forEach(r => { r.img = i + 1; all.rows.push(r); });
            all.notes.push(best.error ? `画像${i + 1}：${best.error}` : `画像${i + 1}：文字${best.wordCount}個・シフト行${best.rows.length}行${best.headerFound ? '' : '（日付の行が読めなかったため、左端の列を' + ocrStartDay + '日として推定）'}`);
            all.rawText += `【画像${i + 1}】\n${best.rawText || ""}\n`;
        }
        const seen = new Set();
        all.rows.forEach(r => { if (r.staff) { const k = r.staff + "|" + r.img; if (seen.has(k)) r.staff = ""; else seen.add(k); } });
        ocrState = { rows: all.rows, rawText: all.rawText, notes: all.notes, error: all.rows.length ? "" : "シフトの行を読み取れませんでした。" };
        if (prog) prog.innerHTML = "";
        renderOcrCard();
        const c = ocrCompare();
        if (c && c.rows.length) toast(`照合しました：一致${c.match}・食い違い${c.diff}・読取不可${c.unread}`, c.diff === 0);
    } catch (err) {
        console.error(err);
        if (prog) prog.innerHTML = `<div class="text-xs font-bold text-red-600">${esc(err.message || "画像認識でエラーが発生しました")}</div>`;
    } finally {
        if (btn) { btn.disabled = false; btn.classList.remove('opacity-60'); }
    }
};

// pdf.js も Worker が使えない環境では、ページ上で処理させる
const _pdfToCanvasesP7 = pdfToCanvases;
pdfToCanvases = async function (file) {
    try { return await _pdfToCanvasesP7(file); }
    catch (e) {
        console.warn("PDF: Worker無しで再試行", e);
        await loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js", "pdfjsLib");
        await loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js", "pdfjsWorker");
        const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()), disableWorker: true, isEvalSupported: false }).promise;
        const out = [];
        for (let p = 1; p <= Math.min(pdf.numPages, 6); p++) {
            const page = await pdf.getPage(p), vp = page.getViewport({ scale: 2.5 });
            const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
            await page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
            out.push(cv);
        }
        return out;
    }
};
</script>
<script>
// ===== 出力3：添付形式の鍵管理表 ＋ 右側にシフト状況（確認用・色分け） =====
function exportToCheckStyle() {
    if (!shiftResult || shiftResult.length === 0) {
        alert("出力するシフトデータがありません。先にシフトを作成してください。");
        return;
    }
    const staff = orderedStaff().map(s => s.name);
    const v = shiftVersions[activeVer];
    const KEY = [
        { k: "鍵No.1", label: "鍵No.１<br>（キーホルダー赤）", head: "#fecaca", soft: "#fef2f2", strong: "#fca5a5", text: "#991b1b" },
        { k: "鍵No.2", label: "鍵No.2<br>（キーホルダー黄）", head: "#fde68a", soft: "#fffbeb", strong: "#fcd34d", text: "#92400e" },
        { k: "鍵No.3", label: "鍵No.3<br>（キーホルダー青）", head: "#bfdbfe", soft: "#eff6ff", strong: "#93c5fd", text: "#1e40af" }
    ];
    const CAT = {
        key:   { name: "鍵の割り当て",       bg: "#ffe4e6", fg: "#be123c" },
        shift: { name: "鍵開け・鍵閉めのシフト不足", bg: "#ede9fe", fg: "#6d28d9" },
        safe:  { name: "金庫チェック",       bg: "#fef3c7", fg: "#b45309" },
        rule:  { name: "担当可・個別ルール", bg: "#e0f2fe", fg: "#0369a1" },
        unset: { name: "担当者未設定",       bg: "#fecaca", fg: "#b91c1c" }
    };
    const shiftStyle = val => {
        const s = String(val || "公休").trim();
        if (!isWorkingShift(s)) return "background:#f1f5f9;color:#94a3b8;";
        if (s.includes("超早")) return "background:#93c5fd;color:#1e3a8a;font-weight:bold;";
        if (s.includes("早") || isEarlyByTime(s)) return "background:#dbeafe;color:#1e40af;font-weight:bold;";
        if (isFullShift(s)) return "background:#ddd6fe;color:#5b21b6;font-weight:bold;";
        if (isLateShift(s)) return "background:#fde68a;color:#92400e;font-weight:bold;";
        return "background:#ffffff;color:#334155;";
    };
    const dowColor = d => d === "日" ? "#dc2626" : d === "土" ? "#2563eb" : "#0f172a";
    const nKeyCols = 12, nRight = staff.length + 5;
    const totalCols = 1 + nKeyCols + 1 + nRight;   // A列 + 鍵12列 + 区切り + 右側
    // 列の幅（px）：文字がセルに収まるよう広めに。Excelでも同じ幅で開くよう pt でも指定
    const colPx = [150, ...new Array(nKeyCols).fill(82), 28, ...staff.map(() => 115), 150, 150, 150, 95, 660];
    const colgroupHtml = colPx.map(px => `<col width="${px}" style="mso-width-source:userset;width:${Math.round(px * 0.75)}pt;">`).join('');

    let html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset="UTF-8">
    <style>
        table { border-collapse: collapse; font-family: 'HG創英角ｺﾞｼｯｸUB', 'Meiryo', sans-serif; }
        th, td { border: 1px solid #000000; padding: 4px 6px; text-align: center; vertical-align: middle; }
        .col-a { width: 160px; font-size: 20pt; font-family: 'MS PGothic', sans-serif; }
        .red-title { background-color: #ff0000; color: #ffffff; font-weight: bold; font-size: 36pt; text-align: left; font-family: 'HG創英角ｺﾞｼｯｸUB', sans-serif; }
        .note-text { text-align: left; font-size: 20pt; font-family: 'HG創英角ｺﾞｼｯｸUB', sans-serif; }
        .header-main { font-weight: bold; font-size: 18pt; font-family: 'HG創英角ｺﾞｼｯｸUB', sans-serif; }
        .sub-header { font-weight: bold; font-size: 18pt; font-family: 'HG創英角ｺﾞｼｯｸUB', sans-serif; }
        .data-cell { font-size: 24pt; font-family: 'MS PGothic', sans-serif; }
        .gap { border: none; width: 24px; background: #ffffff; }
        .r-head { font-weight: bold; font-size: 16pt; font-family: 'Meiryo', sans-serif; }
        .r-cell { font-size: 16pt; font-family: 'Meiryo', sans-serif; white-space: normal; line-height: 1.25; }
        .r-note { font-size: 13pt; font-family: 'Meiryo', sans-serif; text-align: left; white-space: normal; line-height: 1.5; vertical-align: middle; }
        td { mso-number-format: "\@"; }
        .note-text, .r-head { white-space: normal; }
    </style></head>
    <body>
    <table>
        ${colgroupHtml}
        <tr style="height:96pt;"><td colspan="${1 + nKeyCols}" class="red-title">${detectedExcelTitle}</td><td class="gap"></td>
            <td colspan="${nRight}" style="background:#1e293b;color:#ffffff;font-weight:bold;font-size:24pt;text-align:left;font-family:Meiryo;">シフト状況・担当の確認${v ? `（${v.label} 割り当て結果）` : ''}</td></tr>
        <tr style="height:62pt;"><td colspan="${1 + nKeyCols}" class="note-text">＊防犯管理上、いつ、誰が、どの鍵を持っているのか明確にし、受渡し時には必ず両者（引渡者・引受者）が押印して下さい。</td><td class="gap"></td>
            <td colspan="${nRight}" class="r-note" style="font-size:14pt;">色の見方：<span style="background:#93c5fd;">　超早　</span> <span style="background:#dbeafe;">　早　</span> <span style="background:#fde68a;">　遅　</span> <span style="background:#ddd6fe;">　フル　</span> <span style="background:#f1f5f9;color:#94a3b8;">　休・応援・会議など　</span></td></tr>
        <tr style="height:62pt;"><td colspan="${1 + nKeyCols}" class="note-text">＊金庫の鍵は、必ず営業中は携行し、退勤時は持ち帰る運用として下さい。</td><td class="gap"></td>
            <td colspan="${nRight}" class="r-note" style="font-size:14pt;">鍵の欄：色の濃いマス＝その日に鍵を受け取る人（引受者）／ 判定：<span style="background:#dcfce7;color:#15803d;font-weight:bold;">　OK　</span> ${Object.values(CAT).map(c => `<span style="background:${c.bg};color:${c.fg};font-weight:bold;">　${c.name}　</span>`).join(' ')}</td></tr>
        <tr><td colspan="${1 + nKeyCols}" style="border: 1px solid #000;">&nbsp;</td><td class="gap"></td><td colspan="${nRight}" style="border:none;">&nbsp;</td></tr>
        <tr class="header-main" style="height:72pt;">
            <td rowspan="2" class="col-a" style="background:#f2f2f2;">鍵の<br>種類<br>・No.</td>
            ${KEY.map(k => `<td colspan="4" style="background:${k.head};color:${k.text};">${k.label}</td>`).join('')}
            <td class="gap" rowspan="2"></td>
            <td colspan="${staff.length}" class="r-head" style="background:#334155;color:#ffffff;">シフト状況</td>
            <td rowspan="2" class="r-head" style="background:#dbeafe;color:#1e40af;">鍵開け</td>
            <td rowspan="2" class="r-head" style="background:#fde68a;color:#92400e;">鍵閉め</td>
            <td rowspan="2" class="r-head" style="background:#fcd34d;color:#78350f;">金庫<br>チェック</td>
            <td rowspan="2" class="r-head" style="background:#e2e8f0;">判定</td>
            <td rowspan="2" class="r-head" style="background:#e2e8f0;min-width:420px;">NGの内容・確認ポイント</td>
        </tr>
        <tr class="sub-header" style="height:40pt;">
            ${KEY.map(k => `<td colspan="2" style="background:${k.soft};">引渡者</td><td colspan="2" style="background:${k.soft};">引受者</td>`).join('')}
            ${staff.map(n => `<td class="r-head" style="background:#f8fafc;min-width:90px;">${esc(n)}</td>`).join('')}
        </tr>`;

    let prev = { "鍵No.1": "", "鍵No.2": "", "鍵No.3": "" };
    let ngCount = 0;
    shiftResult.forEach((r, idx) => {
        if (!r) return;
        const kh = r.keyHolders || {};
        const pk = idx === 0 ? (r.prevKeyHolders || kh) : prev;
        const shifts = r.shiftsByStaff || {};
        const disp = getDisplayStatus(r);
        const isNG = disp.status === "NG";
        if (isNG) ngCount++;
        const cellMap = isNG ? ngCellMap(r) : {};
        const ngFill = t => cellMap[t] ? `background:${CAT[cellMap[t]].bg};color:${CAT[cellMap[t]].fg};font-weight:bold;` : "";
        const dayNum = String(r.date).split("/").pop();
        const rowBg = isNG ? "#fff1f2" : (idx % 2 ? "#ffffff" : "#fafafa");

        const keyCells = KEY.map((k, ki) => {
            const a = pk[k.k] || "", b = kh[k.k] || "";
            const recv = a === b ? "" : b;
            const ng = ngFill("k" + (ki + 1));
            const store = x => x === STORE_NAME ? "font-style:italic;color:#64748b;" : "";
            return `<td colspan="2" class="data-cell" style="background:${k.soft};${store(a)}${ng}">${esc(a)}</td>` +
                   `<td colspan="2" class="data-cell" style="${recv ? `background:${k.strong};color:${k.text};font-weight:bold;` : `background:${rowBg};`}${store(recv)}${ng}">${esc(recv)}</td>`;
        }).join('');

        const shiftCells = staff.map(n => `<td class="r-cell" style="${shiftStyle(shifts[n])}">${esc(shifts[n] || "公休")}</td>`).join('');
        const who = (name, base, t) => {
            const sv = shifts[name];
            const txt = !name || name === UNSET ? "担当なし" : `${name}${sv ? `<br><span style="font-size:11pt;">(${esc(sv)})</span>` : ''}`;
            return `<td class="r-cell" style="${base}font-weight:bold;${ngFill(t)}">${txt}</td>`;
        };
        const msgs = isNG ? ngMsgs(r) : [];
        const noteHtml = isNG
            ? msgs.map(m => { const c = CAT[ngCat(m)]; return `<span style="background:${c.bg};color:${c.fg};font-weight:bold;">【${c.name}】${esc(m)}</span>`; }).join('<br>')
            : `<span style="color:#475569;">${esc(disp.note || "")}</span>`;

        const noteLines = msgs.reduce((t, m) => t + Math.ceil((m.length + 10) / 36), 0);   // 折り返し後の行数の目安
        const rowH = Math.max(56, noteLines * 21 + 16);
        html += `
        <tr style="height:${rowH}pt;">
            <td class="col-a" style="background:${isNG ? '#fecaca' : '#ffffff'};color:${dowColor(r.dayOfWeek)};font-weight:bold;">${dayNum}日<span style="font-size:14pt;">(${esc(r.dayOfWeek)})</span></td>
            ${keyCells}
            <td class="gap"></td>
            ${shiftCells}
            ${who(r.openRaw, "background:#eff6ff;color:#1e40af;", "open")}
            ${who(r.closeRaw, "background:#fffbeb;color:#92400e;", "close")}
            ${who(r.safeRaw, "background:#fef9c3;color:#78350f;", "safe")}
            <td class="r-cell" style="font-weight:bold;font-size:18pt;${isNG ? 'background:#fee2e2;color:#dc2626;' : 'background:#dcfce7;color:#15803d;'}">${disp.status}</td>
            <td class="r-note" style="background:${rowBg};">${noteHtml}</td>
        </tr>`;
        prev = { "鍵No.1": kh["鍵No.1"] || "", "鍵No.2": kh["鍵No.2"] || "", "鍵No.3": kh["鍵No.3"] || "" };
    });

    // 集計（右下）
    const stats = orderedStaff().map(s => `<td class="r-cell" style="background:#f8fafc;">鍵開け${s.openCount}<br>鍵閉め${s.closeCount}<br>金庫${s.safeCount}</td>`).join('');
    html += `
        <tr><td colspan="${1 + nKeyCols}" style="border:none;">&nbsp;</td><td class="gap"></td><td colspan="${nRight}" style="border:none;">&nbsp;</td></tr>
        <tr style="height:70pt;"><td colspan="${1 + nKeyCols}" style="border:none;"></td><td class="gap"></td>
            ${stats}
            <td colspan="3" class="r-cell" style="background:#f1f5f9;font-weight:bold;">担当回数（左の各人）</td>
            <td class="r-cell" style="font-weight:bold;${ngCount ? 'background:#fee2e2;color:#dc2626;' : 'background:#dcfce7;color:#15803d;'}">NG<br>${ngCount}日</td>
            <td class="r-note" style="background:#f8fafc;">${ngCount ? 'NGの日は日付の欄が赤くなっています。内容を確認してください。' : '全日程OKです。'}</td></tr>
    </table></body></html>`;

    const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const mon = String((shiftResult[0] && String(shiftResult[0].date).split("/")[0]) || detectedMonth || "");
    a.download = `${mon}月鍵管理表予定表_シフト確認用${v ? '_' + v.label : ''}.xls`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    try { toast("確認用の鍵管理表を出力しました", true); } catch (e) { }
}
</script>
<script>
// ===== 手動修正 → そのままExcel出力へ =====
function applyAndGoExport() {
    if (!rawShiftData || !rawShiftData.length) { alert("先にExcelファイルを読み込んでください。"); return; }
    if (!shiftResult || !shiftResult.length) { generateShift(); }
    else {
        const before = ngScore(shiftResult).days;
        _pendingKind = shiftVersions.length ? { kind: "手動修正・再計算", changes: ["画面4での手動修正（固定した日）を反映して再計算"], ngBefore: before } : null;
        fixMessage = "";
        generateShift();
    }
    const v = shiftVersions[activeVer];
    if (v) v.popPending = false;          // 出力画面ではポップアップを出さない
    switchTab('tab5');
    const ng = ngScore(shiftResult).days;
    toast(`${v ? v.label + ' 割り当て結果' : '割り当て結果'}を反映しました${ng ? `（NG ${ng}日あり）` : '（NGなし）'}`, ng === 0);
}
function renderExportStatus() {
    const el = document.getElementById('exportStatus'); if (!el) return;
    if (!shiftResult || !shiftResult.length) { el.innerHTML = ""; return; }
    const v = shiftVersions[activeVer], ng = ngScore(shiftResult).days;
    el.innerHTML = `<div class="p-4 rounded-xl border ${ng ? 'border-amber-300 bg-amber-50' : 'border-emerald-300 bg-emerald-50'} flex flex-wrap items-center gap-3 text-sm">
        <span class="font-bold ${ng ? 'text-amber-800' : 'text-emerald-800'}"><i class="fa-solid ${ng ? 'fa-triangle-exclamation' : 'fa-circle-check'}"></i> 出力する内容：${v ? esc(v.label) + ' 割り当て結果（' + esc(v.kind) + '）' : '割り当て結果'} ／ ${ng ? `NGが${ng}日あります` : 'NGなし'}</span>
        ${ng ? `<button type="button" onclick="switchTab('tab3')" class="text-xs px-3 py-1 rounded-md border border-amber-400 bg-white text-amber-800 hover:bg-amber-100 cursor-pointer">NGを確認する</button><button type="button" onclick="goManualFix()" class="text-xs px-3 py-1 rounded-md border border-amber-400 bg-white text-amber-800 hover:bg-amber-100 cursor-pointer">手動修正に戻る</button>` : ''}
        <span class="text-xs text-slate-500">おすすめ：「確認用」で内容を確かめてから「添付形式」を提出用に使ってください。</span>
      </div>`;
}
const _switchTabExp = switchTab;
switchTab = function (id) { _switchTabExp(id); if (id === 'tab5') renderExportStatus(); };
</script>
<script>
// ===== ブラウザの確認ダイアログ（confirm/alert）が使えない画面（claude.ai のプレビュー等）でも動くように =====
// 確認はページ上のポップアップで行う
function uiConfirm(msg, okLabel, danger) {
    return new Promise(resolve => {
        const old = document.getElementById('uiConfirm'); if (old) old.remove();
        document.body.insertAdjacentHTML('beforeend', `
          <div id="uiConfirm" class="pop-overlay fixed inset-0 z-[80] bg-slate-900/45 flex items-center justify-center p-4">
            <div class="pop-card bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden">
              <div class="p-5 flex gap-3">
                <div class="w-10 h-10 shrink-0 rounded-full ${danger ? 'bg-red-100 text-red-600' : 'bg-indigo-100 text-indigo-600'} flex items-center justify-center"><i class="fa-solid ${danger ? 'fa-rotate-left' : 'fa-circle-question'}"></i></div>
                <div class="text-sm text-slate-700 whitespace-pre-line leading-relaxed">${esc(msg)}</div>
              </div>
              <div class="px-5 py-3 bg-slate-50 border-t border-slate-200 flex justify-end gap-2">
                <button type="button" data-v="0" class="px-4 py-2 rounded-lg border border-slate-300 bg-white text-slate-600 hover:bg-slate-100 text-sm font-semibold cursor-pointer">キャンセル</button>
                <button type="button" data-v="1" class="px-4 py-2 rounded-lg ${danger ? 'bg-red-600 hover:bg-red-700' : 'bg-indigo-600 hover:bg-indigo-700'} text-white text-sm font-bold cursor-pointer">${esc(okLabel || 'OK')}</button>
              </div>
            </div>
          </div>`);
        const box = document.getElementById('uiConfirm');
        box.querySelectorAll('button').forEach(b => b.addEventListener('click', () => { box.remove(); resolve(b.dataset.v === "1"); }));
        box.querySelector('button[data-v="1"]').focus();
    });
}
// confirm を使う機能を「ページ上で確認 → OKなら実行」に包み直す
function withUiConfirm(name, okLabel, danger) {
    const fn = window[name];
    if (typeof fn !== 'function') return;
    window[name] = function (...args) {
        const STOP = { stop: true };
        const realConfirm = window.confirm;
        let asked = null;
        window.confirm = m => { asked = m; throw STOP; };
        try { fn.apply(this, args); }                          // 確認の手前まで動かして、確認文だけ受け取る
        catch (e) { if (e !== STOP) { window.confirm = realConfirm; throw e; } }
        finally { window.confirm = realConfirm; }
        if (asked === null) return;                            // 確認なしで終わった（入力不足など）
        uiConfirm(asked, okLabel, danger).then(ok => {
            if (!ok) return;
            const rc = window.confirm;
            window.confirm = () => true;
            try { fn.apply(this, args); } finally { window.confirm = rc; }
        });
    };
}
withUiConfirm('resetAll', 'リセットする', true);
withUiConfirm('removeStaff', '削除する', true);
withUiConfirm('s0ResetEdits', '元に戻す', true);
withUiConfirm('ocrApplyAll', '画像の値に直す', false);

// alert もページ上のお知らせで表示（ブロックされて何も出ない、を防ぐ）
window.alert = function (msg) {
    const old = document.getElementById('uiAlert'); if (old) old.remove();
    document.body.insertAdjacentHTML('beforeend', `
      <div id="uiAlert" class="pop-overlay fixed inset-0 z-[80] bg-slate-900/40 flex items-center justify-center p-4" onclick="if(event.target===this) this.remove()">
        <div class="pop-card bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden">
          <div class="p-5 flex gap-3"><div class="w-10 h-10 shrink-0 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center"><i class="fa-solid fa-circle-info"></i></div>
            <div class="text-sm text-slate-700 whitespace-pre-line leading-relaxed">${esc(msg)}</div></div>
          <div class="px-5 py-3 bg-slate-50 border-t border-slate-200 flex justify-end"><button type="button" onclick="document.getElementById('uiAlert').remove()" class="px-5 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-bold cursor-pointer">OK</button></div>
        </div>
      </div>`);
};

// リセット後のお知らせ
const _resetAllToast = resetAll;
resetAll = function () {
    _resetAllToast.apply(this, arguments);
    const watch = setInterval(() => {
        if (!document.getElementById('uiConfirm')) {
            clearInterval(watch);
            if (!step0State && (!rawShiftData || !rawShiftData.length)) { try { toast("リセットしました。シフト表・割り当て結果をすべて取り除きました", true); } catch (e) { } }
        }
    }, 200);
};
</script>
<script>
// ===== 画面0：読み取りの注意と、おすすめの進め方（目立つ案内） =====
function aiPromptText() {
    const ym = (typeof getSelectedYM === 'function') ? getSelectedYM() : {};
    const st = (typeof step0State !== 'undefined' && step0State) ? step0State : null;
    const y = (st && st.y) || ym.y || "○○○○", m = (st && st.m) || ym.m || "○";
    const last = (st && st.lastDay) || (ym.y && ym.m ? new Date(ym.y, ym.m, 0).getDate() : "月末");
    const people = orderedStaff().map(s => (s.aliases && s.aliases[0]) ? `${s.aliases[0]}（${s.name}）` : s.name).join("、");
    return `添付したシフト表のExcelを、次の形の新しいExcelファイル（.xlsx）に作り直してください。

【作ってほしい形】
・1行目：A1セルに「${y}年${m}月 シフト一覧」
・2行目：A2セルに「担当者名」、B2セルから右へ 1, 2, 3 … ${last} と日付を数字だけで並べる
・3行目以降：A列に氏名、各日付の列にその日のシフトを1マスずつ入れる
・シフトは元の表記のまま（早・超早・遅・フル・公休・有給・応援・会議・店長会議・研修・時刻表記 など）
・休み・空欄は「公休」と入れる
・色・罫線・結合セル・合計欄・余計な行や列は不要

【対象の人】${people}

作成したら、元の表と1人ずつ・1日ずつ見比べて、写し間違いがないか確認し、違いがあれば直してからファイルを渡してください。`;
}
function copyAiPrompt() {
    const ta = document.getElementById('aiPromptBox');
    const text = aiPromptText();
    if (ta) ta.value = text;
    const done = () => toast("AIへの指示文をコピーしました。ChatGPTなどに元のExcelと一緒に貼り付けてください", true);
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(done).catch(() => { if (ta) { ta.select(); document.execCommand('copy'); } done(); });
    else { if (ta) { ta.select(); document.execCommand('copy'); } done(); }
}
function renderStep0Guide() {
    const el = document.getElementById('step0Guide'); if (!el) return;
    el.innerHTML = `
      <div class="rounded-2xl border-4 border-amber-400 bg-amber-50 shadow-md overflow-hidden">
        <div class="px-5 py-3 bg-amber-400 text-amber-950 flex items-center gap-3">
          <i class="fa-solid fa-triangle-exclamation text-2xl"></i>
          <div class="text-lg sm:text-xl font-black">重要：元のシフト表は、正しく読み取れないことが多いです</div>
        </div>
        <div class="p-5 space-y-4">
          <p class="text-sm sm:text-base text-amber-950 font-bold leading-relaxed">
            店舗のシフト表は、セルの結合・色・時刻表記・レイアウトの違いなどにより、この「0. シフト編集」では<span class="underline decoration-red-500 decoration-4">シフトがしっかり反映されない可能性が高い</span>です。次のどちらかの方法で進めてください。
          </p>
          <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <div class="rounded-xl border-2 border-emerald-500 bg-white p-4 space-y-2 relative">
              <span class="absolute -top-3 left-4 px-3 py-0.5 rounded-full bg-emerald-600 text-white text-sm font-black shadow">おすすめ（推奨）</span>
              <div class="pt-1 text-base font-black text-emerald-800"><i class="fa-solid fa-robot"></i> 方法A：ChatGPTなどのAIでExcelを簡単な形にしてから「1. シフト取込」へ</div>
              <ol class="text-sm text-slate-700 list-decimal list-inside space-y-1">
                <li>下の「指示文をコピー」を押す</li>
                <li>ChatGPT・Claudeなどに<strong>元のシフト表Excel</strong>を添付し、指示文を貼り付けて送る</li>
                <li>できたExcelを<strong>目で元の表と見比べて</strong>確認する</li>
                <li>「1. シフト取込」でそのExcelを読み込む（0は使わなくてOK）</li>
              </ol>
              <div class="flex flex-wrap gap-2 pt-1">
                <button type="button" onclick="copyAiPrompt()" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-4 py-2 rounded-lg text-sm cursor-pointer"><i class="fa-solid fa-copy"></i> AIへの指示文をコピー</button>
                <button type="button" onclick="switchTab('tab1')" class="bg-white border-2 border-emerald-600 text-emerald-700 hover:bg-emerald-50 font-bold px-4 py-2 rounded-lg text-sm cursor-pointer">1. シフト取込へ進む <i class="fa-solid fa-arrow-right"></i></button>
              </div>
              <details class="text-xs"><summary class="cursor-pointer text-slate-500">指示文の中身を見る</summary>
                <textarea id="aiPromptBox" readonly class="mt-2 w-full h-48 p-2 border border-slate-300 rounded text-xs text-slate-700 bg-slate-50">${esc(aiPromptText())}</textarea>
              </details>
            </div>
            <div class="rounded-xl border-2 border-indigo-400 bg-white p-4 space-y-2">
              <div class="text-base font-black text-indigo-800"><i class="fa-solid fa-pen"></i> 方法B：ここで変換したあと、手で直す前提で使う</div>
              <ol class="text-sm text-slate-700 list-decimal list-inside space-y-1">
                <li>下で元のシフト表Excelを選ぶ</li>
                <li>変換された表を<strong>元の表と1マスずつ見比べる</strong></li>
                <li>違うマスはクリックして<strong>手で入力して直す</strong>（黄色になります）</li>
                <li>必要なら一番下の「画像で二重チェック」も使う</li>
              </ol>
              <p class="text-xs text-red-600 font-bold">※変換結果をそのまま使うと、鍵当番・金庫チェックの割り当てが間違うおそれがあります。</p>
            </div>
          </div>
        </div>
      </div>`;
}
(function () {
    const sec = document.getElementById('tab0');
    if (sec) { sec.insertAdjacentHTML('afterbegin', '<div id="step0Guide"></div>'); renderStep0Guide(); }
    const _sw = switchTab;
    switchTab = function (id) { _sw(id); if (id === 'tab0') renderStep0Guide(); };
})();
</script>
<script>
// ===== 手動修正：鍵は「引渡者 → 引受者」で直す =====
const EDIT_KEYS = ["鍵No.1", "鍵No.2", "鍵No.3"];
function prevHoldersOf(i) {
    const it = shiftResult[i]; if (!it) return {};
    return i === 0 ? (it.prevKeyHolders || it.keyHolders || {}) : (shiftResult[i - 1].keyHolders || {});
}
function editKeyCells(i) {
    const it = shiftResult[i]; if (!it) return "";
    const kh = it.keyHolders || {}, pk = prevHoldersOf(i), sh = it.shiftsByStaff || {};
    const names = orderedStaff().map(s => s.name);
    const at = n => n === STORE_NAME || isWorkingShift(sh[n] || "公休");
    const sel = "border border-slate-300 rounded p-1 text-xs bg-white text-center max-w-[7.5rem]";
    return EDIT_KEYS.map(k => {
        const giver = pk[k] || "", now = kh[k] || "", recv = now === giver ? "" : now;
        const giverHtml = i === 0
            ? `<select onchange="setStartHolder('${k}', this.value)" title="月初（前月末）に鍵を持っている人" class="${sel}">${[...names, STORE_NAME].map(n => `<option value="${esc(n)}" ${n === giver ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select><div class="text-[10px] text-slate-400">月初の保持者</div>`
            : `<select onchange="setGiver(${i}, '${k}', this.value)" title="前日の夜に鍵を持っている人（変えると前日の引受者も直ります）" class="${sel} bg-slate-50">${[...names, STORE_NAME].map(n => `<option value="${esc(n)}" ${n === giver ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select>${giver && giver !== STORE_NAME && !at(giver) ? `<div class="text-[10px] text-red-500 font-bold">不在(${esc(sh[giver] || '公休')})</div>` : ''}`;
        const opts = [`<option value="" ${recv === '' ? 'selected' : ''}>―（なし）</option>`]
            .concat(names.filter(n => n !== giver).map(n => {
                const ok = at(n);
                return `<option value="${esc(n)}" ${n === recv ? 'selected' : ''} ${ok || n === recv ? '' : 'disabled'}>${esc(n)} ${esc(sh[n] || '公休')}${ok ? '' : ' ×'}</option>`;
            }))
            .concat(giver === STORE_NAME ? [] : [`<option value="${STORE_NAME}" ${recv === STORE_NAME ? 'selected' : ''}>店舗</option>`]);
        return `<td class="px-2 py-2 text-center whitespace-nowrap border-l-2 border-slate-300">${giverHtml}</td>
                <td class="px-2 py-2 text-center whitespace-nowrap"><select onchange="setReceiver(${i}, '${k}', this.value)" class="${sel} ${recv ? 'font-bold border-2 border-indigo-500 bg-indigo-50' : 'text-slate-400'}">${opts.join('')}</select></td>`;
    }).join('');
}
// 鍵の持ち主を old → nw に変えたら、そのあと「受け渡しなし」で続いている日も一緒に変える
function carryHolder(from, k, old, nw) {
    for (let j = from; j < shiftResult.length; j++) {
        const r = shiftResult[j]; if (!r || r.locked) break;
        if ((r.keyHolders || {})[k] !== old) break;
        r.keyHolders[k] = nw;
    }
}
function afterKeyEdit() {
    try { buildShiftIndex(); validateAll(shiftResult); } catch (e) { }
    renderEditTable();
}
function setReceiver(i, k, v) {
    const it = shiftResult[i]; if (!it) return;
    const giver = prevHoldersOf(i)[k] || "";
    const old = it.keyHolders[k];
    const nw = v || giver;
    if (old === nw) return;
    it.keyHolders[k] = nw;
    it.locked = true;
    carryHolder(i + 1, k, old, nw);
    afterKeyEdit();
    try { toast(v ? `${it.date}：${k} を ${giver} → ${v} に受け渡し` : `${it.date}：${k} は ${giver} が持ち続ける（受け渡しなし）`, true); } catch (e) { }
}
function setStartHolder(k, v) {
    const it = shiftResult[0]; if (!it) return;
    it.prevKeyHolders = it.prevKeyHolders || { ...it.keyHolders };
    const old = it.prevKeyHolders[k];
    if (old === v) return;
    it.prevKeyHolders[k] = v;
    carryHolder(0, k, old, v);
    afterKeyEdit();
}
</script>
<script>
// 引渡者を変える ＝ 前日の夜に鍵を持っている人を変える（前日の引受者として記録）
function setGiver(i, k, v) {
    if (i === 0) { setStartHolder(k, v); return; }
    const prevRow = shiftResult[i - 1], it = shiftResult[i];
    if (!prevRow || !it) return;
    const old = prevRow.keyHolders[k];
    if (old === v) return;
    const noHandover = it.keyHolders[k] === old;        // その日が「受け渡しなし」なら、新しい引渡者が持ち続ける
    prevRow.keyHolders[k] = v;
    prevRow.locked = true;
    if (noHandover) carryHolder(i, k, old, v);
    afterKeyEdit();
    try { toast(`${prevRow.date} の夜は ${k} を ${v} が持っている形にしました（${prevRow.date} の引受者も ${v} に変わります）`, true); } catch (e) { }
}
</script>
<style>
/* ===== 手動修正の表：横スクロールなしで収まるように詰める ===== */
#tab4 table { width: 100%; }
#tab4 #editTableHead th, #tab4 #editTableBody td { padding: 3px 2px !important; }
#tab4 #editTableHead th { font-size: 11px; white-space: normal !important; line-height: 1.2; }
#tab4 #editTableBody select { width: 4.6rem !important; max-width: 4.6rem !important; font-size: 11px !important; padding: 2px 1px !important; }
#tab4 #editTableBody td > div.text-\[10px\] { font-size: 9px; }
#tab4 #editTableBody span.inline-block { min-width: 0 !important; padding: 2px 3px !important; font-size: 11px !important; }
#tab4 #editTableBody td:first-child { font-size: 11px; white-space: nowrap; }
#tab4 #editTableBody td:last-child { min-width: 0 !important; max-width: 10rem; white-space: normal !important; }
#tab4 #editTableBody td:last-child span { font-size: 10px !important; padding: 1px 3px !important; }
#tab4 #editTableHead th.min-w-\[180px\] { min-width: 0 !important; }
#tab4.hide-shift .ed-shift { display: none; }
</style>
<script>
// シフト列の表示・非表示（狭い画面では隠して横スクロールをなくす）
let editShowShift = window.innerWidth >= 1100;
function setEditShowShift(v) { editShowShift = !!v; applyEditShiftCols(); }
function applyEditShiftCols() {
    const sec = document.getElementById('tab4'); if (!sec) return;
    sec.classList.toggle('hide-shift', !editShowShift);
    const cb = document.getElementById('editShowShift'); if (cb) cb.checked = editShowShift;
}
(function () {
    const tb = document.getElementById('editToolbar');
    if (tb) tb.insertAdjacentHTML('afterbegin', `<label class="flex items-center gap-1.5 cursor-pointer font-bold text-slate-700"><input type="checkbox" id="editShowShift" ${editShowShift ? 'checked' : ''} onchange="setEditShowShift(this.checked)"> シフトの列を表示する</label><span class="text-slate-300">｜</span>`);
})();
const _renderEditTableCompact = renderEditTable;
renderEditTable = function () {
    _renderEditTableCompact();
    const n = orderedStaff().length;
    const rows = document.querySelectorAll('#editTableHead tr');
    if (rows[0] && rows[0].children[1]) rows[0].children[1].classList.add('ed-shift');                 // 「シフト状況」の見出し
    if (rows[1]) [...rows[1].children].slice(0, n).forEach(th => th.classList.add('ed-shift'));        // 人の名前の見出し
    document.querySelectorAll('#editTableBody tr').forEach(tr => [...tr.children].slice(1, 1 + n).forEach(td => td.classList.add('ed-shift')));
    applyEditShiftCols();
};
</script>
<style>
/* ===== 手動修正：シフトは日付の下にまとめて表示（横スクロールなし） ===== */
#tab4 table.edit-compact { width: 100%; table-layout: auto; }
#tab4 table.edit-compact th, #tab4 table.edit-compact td { padding: 4px 4px !important; vertical-align: middle; }
#tab4 table.edit-compact select { width: 4.8rem !important; max-width: 4.8rem !important; font-size: 11px !important; padding: 2px 1px !important; }
#tab4 .ed-chip { display: inline-block; font-size: 10px; font-weight: 700; padding: 1px 4px; border-radius: 4px; margin: 1px; white-space: nowrap; line-height: 1.4; }
#tab4 .ed-role { display: flex; align-items: center; gap: 3px; margin: 2px 0; border-radius: 6px; padding: 1px 2px; }
#tab4 .ed-role b { font-size: 10px; width: 1.6rem; text-align: right; color: #475569; }
#tab4 .ed-key { display: flex; flex-direction: column; align-items: center; gap: 1px; border-radius: 6px; padding: 2px; }
#tab4 .ed-key .lbl { font-size: 9px; color: #64748b; line-height: 1; }
#tab4 td.ed-judge { max-width: 11rem; white-space: normal !important; text-align: left; }
#tab4 td.ed-judge span { font-size: 10px !important; padding: 1px 3px !important; }
#editShowShiftWrap { display: none !important; }
#tab4 table.edit-compact td:first-child { white-space: normal !important; }
#tab4 .ed-key div { font-size: 9px !important; line-height: 1.15; font-weight: 700; }
#tab4 .ed-key div.text-red-500 { color: #ef4444; }
</style>
<script>
const ED_CHIP = v => {
    const s = String(v || "公休").trim();
    if (!isWorkingShift(s)) return "background:#f1f5f9;color:#94a3b8;";
    if (s.includes("超早")) return "background:#bfdbfe;color:#1e3a8a;";
    if (s.includes("早") || isEarlyByTime(s)) return "background:#e0f2fe;color:#075985;";
    if (isFullShift(s)) return "background:#ede9fe;color:#5b21b6;";
    if (isLateShift(s)) return "background:#fef3c7;color:#92400e;";
    return "background:#f8fafc;color:#334155;";
};
const ED_TINT = { key: "#ffe4e6", shift: "#ede9fe", safe: "#fef3c7", rule: "#e0f2fe", unset: "#fecaca" };

renderEditTable = function () {
    try {
        const body = document.getElementById('editTableBody'), head = document.getElementById('editTableHead');
        if (!body || !head) return;
        const table = body.closest('table'); if (table) table.classList.add('edit-compact');
        if ((!shiftResult || !shiftResult.length) && rawShiftData && rawShiftData.length) generateShift();
        const names = orderedStaff().map(s => s.name);
        const keyHead = (c, b, t) => `<th class="text-center whitespace-nowrap" style="background:${c};border-top:4px solid ${b};font-size:11px;">${t}<div style="font-size:9px;font-weight:normal;color:#64748b;">引渡者 → 引受者</div></th>`;
        head.innerHTML = `<tr>
            <th class="text-center bg-slate-100" style="font-size:11px;">日付・シフト<div style="font-size:9px;font-weight:normal;color:#64748b;">固定</div></th>
            <th class="text-center whitespace-nowrap" style="font-size:11px;background:#e0f2fe;">鍵開け</th>
            <th class="text-center whitespace-nowrap" style="font-size:11px;background:#fef3c7;">鍵閉め</th>
            <th class="text-center whitespace-nowrap" style="font-size:11px;background:#fde68a;">金庫<br>チェック</th>
            ${keyHead('#fef2f2', '#f87171', '鍵No.1 (赤)')}${keyHead('#fffbeb', '#fbbf24', '鍵No.2 (黄)')}${keyHead('#eff6ff', '#60a5fa', '鍵No.3 (青)')}
            <th class="text-center bg-slate-100" style="font-size:11px;">判定（NG理由）</th></tr>`;
        body.innerHTML = "";
        if (!shiftResult || !shiftResult.length) return;
        const ngOnly = (document.getElementById('editNgOnly') || {}).checked;
        const sel = "border border-slate-300 rounded bg-white text-center";

        shiftResult.forEach((it, i) => {
            if (!it) return;
            const sh = it.shiftsByStaff || {};
            const disp = getDisplayStatus(it), isNG = disp.status === "NG";
            const map = isNG ? ngCellMap(it) : {};
            const tint = t => map[t] ? `background:${ED_TINT[map[t]]};outline:2px solid ${ED_TINT[map[t]] === '#fecaca' ? '#f87171' : 'transparent'};` : "";
            const dc = it.dayOfWeek === "日" ? "#dc2626" : it.dayOfWeek === "土" ? "#2563eb" : "#0f172a";

            const chips = names.map(n => `<span class="ed-chip" style="${ED_CHIP(sh[n])}" title="${esc(n)}：${esc(sh[n] || '公休')}">${esc(n)} ${esc(sh[n] || '公休')}</span>`).join('');
            const roleOpts = (cur, safe) => names.map(n => {
                const v = sh[n] || "公休";
                const ok = !safe || !RULES.safeLateOnly || canSafeShift(v);
                return `<option value="${esc(n)}" ${cur === n ? 'selected' : ''} ${ok || cur === n ? '' : 'disabled'}>${esc(n)} ${esc(v)}${ok ? '' : ' ×'}</option>`;
            }).join('') + `<option value="${UNSET}" ${cur === UNSET ? 'selected' : ''}>未設定</option>`;
            const role = (field, cur, t, safe) => `<td class="text-center" style="${tint(t)}"><select onchange="updateShiftField(${i}, '${field}', this.value); afterKeyEdit();" class="${sel}">${roleOpts(cur, safe)}</select></td>`;

            // 鍵：引渡者 → 引受者（既存の部品から選択欄だけ取り出して縦に並べる）
            const tmp = document.createElement('tbody');
            tmp.innerHTML = `<tr>${editKeyCells(i)}</tr>`;
            const cells = [...tmp.querySelectorAll('td')];
            const keyCol = (ki) => {
                const g = cells[ki * 2], r = cells[ki * 2 + 1];
                return `<td class="text-center"><div class="ed-key" style="${tint('k' + (ki + 1))}">
                    <span class="lbl">引渡者</span>${g ? g.innerHTML : ''}
                    <span style="color:#94a3b8;font-size:10px;line-height:1;">↓</span>
                    <span class="lbl">引受者</span>${r ? r.innerHTML : ''}</div></td>`;
            };
            const judge = isNG
                ? `<div style="color:#dc2626;font-weight:800;font-size:12px;">NG ${ngDots(it)}</div><div class="flex flex-wrap gap-0.5 mt-0.5">${ngMsgs(it).map(ngChip).join('')}</div>`
                : `<div style="color:#059669;font-weight:800;font-size:12px;">OK</div><div style="font-size:10px;color:#64748b;">${esc(disp.note || '')}</div>`;

            const tr = document.createElement('tr');
            if (isNG) tr.style.background = "#fff1f2";
            tr.innerHTML = `
                <td style="min-width:10.5rem;max-width:12rem;">
                    <div class="flex items-center justify-between gap-1"><span style="font-weight:800;font-size:12px;color:${dc};">${esc(it.date)}(${esc(it.dayOfWeek)})</span>
                    <label style="font-size:10px;color:#64748b;" class="flex items-center gap-0.5 cursor-pointer"><input type="checkbox" ${it.locked ? 'checked' : ''} onchange="toggleLock(${i}, this.checked)">固定</label></div>
                    <div style="margin-top:2px;">${chips}</div>
                </td>
                ${role('openStaff', it.openRaw, 'open', false)}${role('closeStaff', it.closeRaw, 'close', false)}${role('safeStaff', it.safeRaw, 'safe', true)}
                ${keyCol(0)}${keyCol(1)}${keyCol(2)}
                <td class="ed-judge">${judge}</td>`;
            if (ngOnly && !isNG) tr.classList.add('hidden');
            tr.style.animationDelay = Math.min(i * 12, 350) + 'ms';
            body.appendChild(tr);
        });

        const c = document.getElementById('editNgCount');
        if (c) { const n = ngScore(shiftResult).days; c.innerHTML = n ? `<span class="font-bold text-red-600">NG ${n}日</span>（赤い行）` : '<span class="font-bold text-emerald-600">NGなし</span>'; }
        const lg = document.getElementById('editNgLegend'); if (lg) lg.innerHTML = ngLegend(shiftResult);
    } catch (err) { console.error("renderEditTable error:", err); }
};
// 「シフトの列を表示する」切替は不要になったので隠す
(function () { const cb = document.getElementById('editShowShift'); if (cb && cb.parentElement) { cb.parentElement.id = 'editShowShiftWrap'; const sep = cb.parentElement.nextElementSibling; if (sep) sep.style.display = 'none'; } const t4 = document.getElementById('tab4'); if (t4) t4.classList.remove('hide-shift'); })();
applyEditShiftCols = function () { const t4 = document.getElementById('tab4'); if (t4) t4.classList.remove('hide-shift'); };
</script>
<style>
#tab4 #editTableBody select { width: 4rem !important; max-width: 4rem !important; font-size: 10.5px !important; }
#tab4 table.edit-compact th, #tab4 table.edit-compact td { padding: 3px 2px !important; }
#tab4 tr.ed-ngrow td { padding: 2px 6px 6px !important; border-top: none !important; }
#tab4 .ed-badge { display: inline-block; font-size: 10px; font-weight: 800; padding: 0 5px; border-radius: 9999px; }
</style>
<script>
// ===== 手動修正：鍵は結果画面と同じ「引渡者｜引受者」の横並び。NG理由は日の下の細い行に =====
renderEditTable = function () {
    try {
        const body = document.getElementById('editTableBody'), head = document.getElementById('editTableHead');
        if (!body || !head) return;
        const table = body.closest('table'); if (table) table.classList.add('edit-compact');
        if ((!shiftResult || !shiftResult.length) && rawShiftData && rawShiftData.length) generateShift();
        const names = orderedStaff().map(s => s.name);
        const K = [["鍵No.1 (赤)", "#fef2f2", "#f87171"], ["鍵No.2 (黄)", "#fffbeb", "#fbbf24"], ["鍵No.3 (青)", "#eff6ff", "#60a5fa"]];
        head.innerHTML = `
          <tr>
            <th rowspan="2" class="text-center bg-slate-100" style="font-size:11px;">日付・シフト</th>
            <th rowspan="2" class="text-center whitespace-nowrap" style="font-size:11px;background:#e0f2fe;">鍵開け</th>
            <th rowspan="2" class="text-center whitespace-nowrap" style="font-size:11px;background:#fef3c7;">鍵閉め</th>
            <th rowspan="2" class="text-center" style="font-size:11px;background:#fde68a;">金庫<br>チェック</th>
            ${K.map(([t, bg, bd]) => `<th colspan="2" class="text-center whitespace-nowrap" style="font-size:11px;background:${bg};border-top:4px solid ${bd};border-left:2px solid #cbd5e1;">${t}</th>`).join('')}
          </tr>
          <tr>${K.map(([t, bg]) => `<th class="text-center whitespace-nowrap" style="font-size:10px;background:${bg};color:#475569;border-left:2px solid #cbd5e1;">引渡者</th><th class="text-center whitespace-nowrap" style="font-size:10px;background:${bg};color:#475569;">引受者</th>`).join('')}</tr>`;
        body.innerHTML = "";
        if (!shiftResult || !shiftResult.length) return;
        const ngOnly = (document.getElementById('editNgOnly') || {}).checked;
        const sel = "border border-slate-300 rounded bg-white text-center";

        shiftResult.forEach((it, i) => {
            if (!it) return;
            const sh = it.shiftsByStaff || {};
            const disp = getDisplayStatus(it), isNG = disp.status === "NG";
            const map = isNG ? ngCellMap(it) : {};
            const tint = t => map[t] ? `background:${ED_TINT[map[t]]};` : "";
            const dc = it.dayOfWeek === "日" ? "#dc2626" : it.dayOfWeek === "土" ? "#2563eb" : "#0f172a";
            const chips = names.map(n => `<span class="ed-chip" style="${ED_CHIP(sh[n])}" title="${esc(n)}：${esc(sh[n] || '公休')}">${esc(n)} ${esc(sh[n] || '公休')}</span>`).join('');
            const roleOpts = (cur, safe) => names.map(n => {
                const v = sh[n] || "公休";
                const ok = !safe || !RULES.safeLateOnly || canSafeShift(v);
                return `<option value="${esc(n)}" ${cur === n ? 'selected' : ''} ${ok || cur === n ? '' : 'disabled'}>${esc(n)} ${esc(v)}${ok ? '' : ' ×'}</option>`;
            }).join('') + `<option value="${UNSET}" ${cur === UNSET ? 'selected' : ''}>未設定</option>`;
            const role = (field, cur, t, safe) => `<td class="text-center" style="${tint(t)}"><select onchange="updateShiftField(${i}, '${field}', this.value); afterKeyEdit();" class="${sel}">${roleOpts(cur, safe)}</select></td>`;

            const tmp = document.createElement('tbody');
            tmp.innerHTML = `<tr>${editKeyCells(i)}</tr>`;
            const cells = [...tmp.querySelectorAll('td')];
            const keyTds = [0, 1, 2].map(ki => {
                const g = cells[ki * 2], r = cells[ki * 2 + 1], t = tint('k' + (ki + 1));
                return `<td class="text-center" style="border-left:2px solid #cbd5e1;${t}">${g ? g.innerHTML : ''}</td><td class="text-center" style="${t}">${r ? r.innerHTML : ''}</td>`;
            }).join('');

            const tr = document.createElement('tr');
            tr.dataset.idx = i;
            if (isNG) tr.style.background = "#fff1f2";
            tr.innerHTML = `
                <td style="min-width:8.5rem;max-width:10rem;">
                    <div class="flex items-center gap-1 flex-wrap"><span style="font-weight:800;font-size:12px;color:${dc};">${esc(it.date)}(${esc(it.dayOfWeek)})</span>
                      <span class="ed-badge" style="${isNG ? 'background:#fee2e2;color:#dc2626;' : 'background:#dcfce7;color:#059669;'}">${disp.status}</span>${isNG ? ngDots(it) : ''}
                      <label style="font-size:10px;color:#64748b;margin-left:auto;" class="flex items-center gap-0.5 cursor-pointer"><input type="checkbox" ${it.locked ? 'checked' : ''} onchange="toggleLock(${i}, this.checked)">固定</label></div>
                    <div style="margin-top:2px;">${chips}</div>
                </td>
                ${role('openStaff', it.openRaw, 'open', false)}${role('closeStaff', it.closeRaw, 'close', false)}${role('safeStaff', it.safeRaw, 'safe', true)}
                ${keyTds}`;
            if (ngOnly && !isNG) tr.classList.add('hidden');
            body.appendChild(tr);
            if (false) {
                const sub = document.createElement('tr');
                sub.className = 'ed-ngrow';
                sub.dataset.ngof = i;
                sub.style.background = "#fff1f2";
                sub.innerHTML = `<td colspan="10" style="text-align:left;"><span style="font-size:10px;font-weight:800;color:#dc2626;margin-right:4px;">NG理由：</span><span class="inline-flex flex-wrap gap-1 align-middle">${ngMsgs(it).map(ngChip).join('')}</span></td>`;
                body.appendChild(sub);
            }
        });

        const c = document.getElementById('editNgCount');
        if (c) { const n = ngScore(shiftResult).days; c.innerHTML = n ? `<span class="font-bold text-red-600">NG ${n}日</span>（赤い行）` : '<span class="font-bold text-emerald-600">NGなし</span>'; }
        const lg = document.getElementById('editNgLegend'); if (lg) lg.innerHTML = ngLegend(shiftResult);
    } catch (err) { console.error("renderEditTable error:", err); }
};
// NG理由の行が入ったので、日付の行は番号で探す
goManualFixDay = function (i) {
    goManualFix();
    setTimeout(() => {
        const tr = document.querySelector(`#editTableBody tr[data-idx="${i}"]`);
        if (tr) { tr.scrollIntoView({ behavior: 'smooth', block: 'center' }); tr.classList.add('flash-row'); }
    }, 60);
};
</script>
<script>
// ===== 手動修正の反映：自動で作り直さず、画面の内容をそのまま確定してチェックだけし直す =====
function reflectManualEdits(goExport) {
    if (!shiftResult || !shiftResult.length) { alert("先に割り当てを作成してください。"); return; }
    buildShiftIndex();
    validateAll(shiftResult);
    recount(shiftResult);
    const cur = shiftVersions[activeVer];
    const before = cur ? ngScore(cur.result).days : null;
    pushVersion("手動修正", ["画面4で手動修正した内容を反映（自動での作り直しはしない）"], before);
    const v = shiftVersions[activeVer];
    fixMessage = "";
    if (goExport) { if (v) v.popPending = false; switchTab('tab5'); }
    else switchTab('tab3');
    const ng = ngScore(shiftResult).days;
    try { toast(`${v.label} 割り当て結果として反映しました${ng ? `（NG ${ng}日）` : '（NGなし）'}`, ng === 0); } catch (e) { }
}
applyAndGoExport = function () {
    if (!rawShiftData || !rawShiftData.length) { alert("先にExcelファイルを読み込んでください。"); return; }
    if (!shiftResult || !shiftResult.length) generateShift();
    reflectManualEdits(true);
};

// 手で直した日は「手動修正」として印を付け、あとで自動修正・再計算しても上書きしない
const _updateShiftFieldM = updateShiftField;
updateShiftField = function (index, field, value) { _updateShiftFieldM(index, field, value); const r = shiftResult[index]; if (r) { r.locked = true; r.manualEdited = true; } };
const _setReceiverM = setReceiver;
setReceiver = function (i, k, v) { if (shiftResult[i]) shiftResult[i].manualEdited = true; _setReceiverM(i, k, v); };
const _setGiverM = setGiver;
setGiver = function (i, k, v) { if (i > 0 && shiftResult[i - 1]) shiftResult[i - 1].manualEdited = true; _setGiverM(i, k, v); };
// 鍵を変えたあと「受け渡しなし」で続く日も、変えた内容のまま固定する（作り直しで崩れないように）
carryHolder = function (from, k, old, nw) {
    for (let j = from; j < shiftResult.length; j++) {
        const r = shiftResult[j]; if (!r || r.manualEdited) break;
        if ((r.keyHolders || {})[k] !== old) break;
        r.keyHolders[k] = nw;
        r.locked = true;
    }
};

// ===== 手動修正：NGの一覧は表の上にまとめて表示（クリックでその日へ） =====
function jumpEditRow(i) {
    const tr = document.querySelector(`#editTableBody tr[data-idx="${i}"]`);
    if (tr) { tr.scrollIntoView({ behavior: 'smooth', block: 'center' }); tr.classList.remove('flash-row'); void tr.offsetWidth; tr.classList.add('flash-row'); }
}
const _renderEditTableNgList = renderEditTable;
renderEditTable = function () {
    _renderEditTableNgList();
    const lg = document.getElementById('editNgLegend'); if (!lg || !shiftResult || !shiftResult.length) return;
    const bad = shiftResult.map((r, i) => ({ r, i })).filter(o => o.r && getDisplayStatus(o.r).status === 'NG');
    lg.innerHTML = !bad.length ? '<div class="p-3 rounded-lg border border-emerald-300 bg-emerald-50 text-emerald-800 text-sm font-bold"><i class="fa-solid fa-circle-check"></i> NGはありません</div>' : `
      <div class="rounded-lg border border-red-200 bg-red-50 p-3 space-y-2">
        <div class="flex flex-wrap items-center gap-2"><span class="font-bold text-red-700 text-sm"><i class="fa-solid fa-triangle-exclamation"></i> NGの日（${bad.length}日）</span><span class="text-[11px] text-slate-500">日付を押すと表のその日へ移動します</span></div>
        ${ngLegend(shiftResult)}
        <ul class="space-y-1 max-h-48 overflow-auto pr-1">${bad.map(o => `
          <li class="flex gap-2 items-start text-xs"><button type="button" onclick="jumpEditRow(${o.i})" class="shrink-0 font-bold text-red-700 underline decoration-dotted hover:text-red-900 cursor-pointer whitespace-nowrap">${esc(o.r.date)}(${esc(o.r.dayOfWeek)})</button>
            <span class="flex flex-wrap gap-1">${ngMsgs(o.r).map(ngChip).join('')}</span></li>`).join('')}</ul>
      </div>`;
};
</script>
</body>
</html>
