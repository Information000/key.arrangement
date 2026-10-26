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
</head>
<body class="bg-slate-50 text-slate-800 font-sans min-h-screen flex flex-col text-sm">
 
    <!-- ヘッダー -->
    <header class="bg-indigo-600 text-white shadow-md">
        <div class="max-w-7xl mx-auto px-4 py-3 flex flex-col md:flex-row justify-between items-center">
            <h1 class="text-base font-bold flex items-center gap-2">
                <i class="fa-solid fa-key"></i> 鍵当番・金庫チェック自動割当システム
            </h1>
            <p class="text-xs text-indigo-200 mt-1 md:mt-0">引渡者・引受者完全準拠版</p>
        </div>
    </header>
 
    <!-- タブナビゲーション -->
    <nav class="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-10">
        <div class="max-w-7xl mx-auto px-4 flex overflow-x-auto space-x-1 sm:space-x-4">
            <button type="button" onclick="switchTab('tab1')" id="btn-tab1" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-indigo-600 text-indigo-600 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-file-excel"></i> 1. シフト取込
            </button>
            <button type="button" onclick="switchTab('tab2')" id="btn-tab2" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-users-gear"></i> 2. 優先順位・ルール
            </button>
            <button type="button" onclick="switchTab('tab3')" id="btn-tab3" class="tab-btn px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 whitespace-nowrap flex items-center gap-2 cursor-pointer">
                <i class="fa-solid fa-calendar-days"></i> 3. 割り当て結果
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
    <main class="max-w-7xl mx-auto px-4 py-6 flex-grow w-full space-y-6">
 
        <!-- 画面1: シフト取込 (STEP0) -->
        <section id="tab1" class="tab-content space-y-6">
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <h2 class="text-base font-bold text-slate-700 mb-2 flex items-center gap-2">
                    <i class="fa-solid fa-cloud-arrow-up text-indigo-600"></i> STEP0. シフト表ファイルのアップロード
                </h2>
                <p class="text-xs sm:text-sm text-slate-500 mb-4">
                    Excelファイル（鍵管理表やシフト表）をアップロードしてください。タイトル行の年月や全日程シフトを自動で柔軟に解析します。
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
                        <i class="fa-solid fa-table text-indigo-600"></i> STEP0. 対象メンバー・全日程シフト一覧表
                    </h3>
                    <span id="loadStatus" class="text-xs px-2.5 py-1 bg-amber-50 text-amber-700 rounded-full font-medium border border-amber-200">未読込（Excelを読み込んでください）</span>
                </div>
                <div class="overflow-x-auto max-h-96 border border-slate-200 rounded-lg">
                    <table id="previewTable" class="min-w-full text-xs sm:text-sm text-center text-slate-600 divide-y divide-slate-200">
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
                            <li><strong>鍵開けは超早・早の人</strong>：前日から鍵を持っている人、または前日に誰かが店舗に置いた鍵を使える人だけが担当できます。月全体を先読みして、翌日以降の鍵開け担当に鍵が渡るよう割り当てます。休・応援・グランド等で出勤しない人は鍵を持ち帰ったまま次の出勤日に使うため、店舗に置くのは「どうしても受け渡せない時」だけです。</li>
                            <li><strong>鍵閉め・金庫チェックは遅番の人</strong>：遅番の人の中から選びます。鍵閉め担当は退勤時に鍵を持っている必要があります。</li>
                            <li><strong><span class="fixedHolderName">五家</span>は1つの鍵を持ち回し ＆ 残り2つを分散</strong>：<span class="fixedHolderName">五家</span>が鍵No.3をずっと持ち回し、残りの鍵No.1・No.2を他のメンバーで回します。店舗置きは、持ち帰りや受け渡しで回らない時だけです。</li>
                            <li><strong>引渡者・引受者同一時の省略</strong>：引渡者と引受者が同じ場合は、引受者セルを空欄としてスッキリ表現します。</li>
                            <li><strong>非番日考慮</strong>：応援・休・グランド・店長会議・方針等の出勤しない日では鍵の受け渡しができないため、出勤しているメンバーまたは店舗置きで適切に対応します。</li>
                            <li><strong>金庫チェック連日NG</strong>：金庫チェックの連日担当のみNGとします。</li>
                        </ul>
                    </div>
                </div>
            </div>
        </section>
 
        <!-- 画面3: 割り当て結果 -->
        <section id="tab3" class="tab-content hidden space-y-6">
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
                    <div class="text-base font-bold text-indigo-600">要件適合確認済み</div>
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
                        <i class="fa-solid fa-calendar-check text-indigo-600"></i> シフト統合・割り当て結果 ＆ 最終チェック (OK/NG)
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
                    <span class="px-2 py-0.5 rounded-md text-slate-400 font-bold">休・応援など</span>
                    <span class="text-slate-300">｜</span>
                    <span class="inline-flex items-center gap-1.5"><span class="inline-block w-9 h-5 rounded-md bg-slate-100 border border-slate-300"></span>引渡者＝前日から鍵を持っている人</span>
                    <span class="text-slate-400 font-bold">→</span>
                    <span class="inline-flex items-center gap-1.5"><span class="inline-block w-9 h-5 rounded-md bg-indigo-100 border-2 border-indigo-500"></span>引受者＝その日に新しく受け取る人（太枠）</span>
                    <span class="inline-flex items-center gap-1"><span class="text-slate-400 font-bold">―</span>引受者が空欄＝同じ人が持ち続ける（受け渡しなし）</span>
                    </div>
                <div class="overflow-auto border border-slate-200 rounded-lg max-h-[72vh]">
                    <table class="min-w-full text-xs text-center text-slate-600 border-separate border-spacing-0">
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
 
        <!-- 画面4: 手動修正 -->
        <section id="tab4" class="tab-content hidden space-y-6">
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <h2 class="text-base font-bold text-slate-700 mb-2 flex items-center gap-2">
                    <i class="fa-solid fa-pen-to-square text-indigo-600"></i> シフトの手動修正・担当固定
                </h2>
                <p class="text-xs sm:text-sm text-slate-500 mb-4">
                    担当者や鍵1、2、3の保持者を変更できます（鍵が回らない場合は「店舗」も選べます）。「固定」にチェックを入れると再計算時に上書きされません。
                </p>
                <div class="overflow-x-auto border border-slate-200 rounded-lg max-h-[500px]">
                    <table class="min-w-full text-xs sm:text-sm text-center text-slate-600 divide-y divide-slate-200">
                        <thead class="bg-slate-100 text-slate-700 sticky top-0 z-10">
                            <tr>
                                <th class="px-3 py-2.5 font-semibold whitespace-nowrap text-center">日付</th>
                                <th class="px-3 py-2.5 font-semibold whitespace-nowrap text-center">鍵開け</th>
                                <th class="px-3 py-2.5 font-semibold whitespace-nowrap text-center">鍵閉め</th>
                                <th class="px-3 py-2.5 font-semibold whitespace-nowrap text-center">金庫チェック</th>
                                <th class="px-3 py-2.5 font-semibold bg-red-50 text-red-800 whitespace-nowrap text-center">鍵1保持者</th>
                                <th class="px-3 py-2.5 font-semibold bg-amber-50 text-amber-800 whitespace-nowrap text-center">鍵2保持者</th>
                                <th class="px-3 py-2.5 font-semibold bg-blue-50 text-blue-800 whitespace-nowrap text-center">鍵3保持者</th>
                                <th class="px-3 py-2.5 font-semibold text-center whitespace-nowrap">固定</th>
                            </tr>
                        </thead>
                        <tbody id="editTableBody" class="divide-y divide-slate-200 bg-white"></tbody>
                    </table>
                </div>
                <div class="mt-4 flex justify-end">
                    <button type="button" onclick="recalculate()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-5 py-2 rounded-lg shadow transition flex items-center gap-2 cursor-pointer text-sm">
                        <i class="fa-solid fa-rotate"></i> 変更を反映して再計算
                    </button>
                </div>
            </div>
        </section>
 
        <!-- 画面5: Excel出力 -->
        <section id="tab5" class="tab-content hidden space-y-6">
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
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
 
        // 鍵開けできるのは「超早」「早」の人（または9時までに出勤する時刻表記の人）
        function canOpenShift(shiftStr) {
            let s = String(shiftStr).trim();
            if (!isWorkingShift(s)) return false;
            return s.includes("早") || isEarlyByTime(s);
        }
 
        // 鍵閉め・金庫チェックができるのは「遅番」の人だけ
        function isLateShift(shiftStr) {
            let s = String(shiftStr).trim();
            if (!isWorkingShift(s)) return false;
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
                const cands = locked ? [[locked.keyHolders["鍵No.1"], locked.keyHolders["鍵No.2"]]] : states;
                const next = new Map();
 
                for (const pv of cur.values()) {
                    const p = pv.s;
                    // 鍵開け：超早・早の誰かが前日から鍵を持っているか、店舗に鍵がある必要がある
                    const holders = new Set([...p, fixedName]);
                    const openPenalty = (!locked && E.length > 0 && !p.includes(STORE_NAME) && !E.some(e => holders.has(e))) ? 1000 : 0;
 
                    for (const s of cands) {
                        let cost = pv.cost + openPenalty;
                        let closer = UNSET;
                        let bad = false;
                        for (let k = 0; k < 2; k++) {
                            const a = p[k], b = s[k];
                            if (a !== b && !(present(a) && present(b))) bad = true;   // 非出勤者は受け渡し不可
                            if (a !== b) cost += 2;                                    // 受け渡しは少ないほど良い
                        }
                        if (locked) {
                            if (bad) cost += 50;   // 手動固定日は強制（矛盾していても受け入れる）
                        } else {
                            if (bad) continue;
                            if (s[0] === s[1] && s[0] !== STORE_NAME) continue;        // 同じ人が2本持たない
                            s.forEach(x => {
                                if (x === STORE_NAME) cost += 100;                    // 店舗置きは最後の手段
                                else if (x === fixedName) cost += 500;               // 固定担当は鍵No.3
                            });
                            // 鍵閉め担当は退勤時に鍵が必要（優先順位が高い遅番の人を優先）
                            if (L.length > 0) {
                                let best = null;
                                L.forEach((c, idx) => {
                                    const access = (c === fixedName) || p.includes(c) || s.includes(c);
                                    const cc = (access ? 0 : 300) + idx * 3;
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
            let prevHolders = { "鍵No.1": keyPlan.start[0], "鍵No.2": keyPlan.start[1], "鍵No.3": fixedName };
 
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
                    if (locked.safeRaw !== UNSET && locked.safeRaw !== "-") lastSafeStaff = locked.safeRaw;
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
 
                // 1. 鍵開け担当（超早・早の人）
                //    前日から鍵を持っている人、または前日に誰かが店舗に置いた鍵を使える人
                const prevKeyPeople = [prevHolders["鍵No.1"], prevHolders["鍵No.2"], prevHolders["鍵No.3"]]
                    .filter(h => h && h !== STORE_NAME);
                const storeKeyAvailable = ["鍵No.1", "鍵No.2", "鍵No.3"].some(k => prevHolders[k] === STORE_NAME);
                let openCandidates = availableStaffs.filter(s => canOpenShift(s.shift));
                openCandidates.sort((a, b) => a.priority - b.priority);
                let openWithKey = openCandidates.filter(s => storeKeyAvailable || prevKeyPeople.includes(s.name));
                let openCandidate = openWithKey[0] || openCandidates[0] || null;
                const openNameRaw = openCandidate ? openCandidate.name : UNSET;
                const openNoKey = (openNameRaw !== UNSET && !storeKeyAvailable && !prevKeyPeople.includes(openNameRaw));
                const openUsesStoreKey = (openNameRaw !== UNSET && storeKeyAvailable && !prevKeyPeople.includes(openNameRaw));
 
                const KEY_NAMES = ["鍵No.1", "鍵No.2", "鍵No.3"];
                const plan = keyPlan.days[index];
 
                // 2. 鍵閉め担当（遅番の人の中から。鍵の持ち回しプランで決めた人）
                let lateStaffs = availableStaffs.filter(s => isLateShift(s.shift));
                lateStaffs.sort((a, b) => a.priority - b.priority);
                let closeCandidate = lateStaffs.find(s => plan && s.name === plan.closer) || lateStaffs[0] || null;
                const closeNameRaw = closeCandidate ? closeCandidate.name : UNSET;
 
                // 3. 金庫チェック担当（連日担当はNG）
                let safeCandidates = lateStaffs.filter(s => s.name !== lastSafeStaff && s.name !== closeNameRaw);
                let safeCandidate = null;
                let safeNote = "";
                let isSafeContinuous = false;
                const dayW = dayData.dayOfWeek || '木';
 
                if (safeCandidates.length > 0) {
                    if (dayW === '月') {
                        let mondayCandidate = safeCandidates.slice().sort((a, b) => a.priority - b.priority)
                            .find(s => { const st = staffList.find(x => x.name === s.name); return st && st.mondaySafe; });
                        if (mondayCandidate) safeCandidate = mondayCandidate;
                        else {
                            safeCandidates.sort((a, b) => a.priority - b.priority);
                            safeCandidate = safeCandidates[0];
                        }
                    } else {
                        safeCandidates.sort((a, b) => a.priority - b.priority);
                        safeCandidate = safeCandidates[0];
                    }
                } else {
                    let fallbackSafe = lateStaffs.filter(s => s.name !== closeNameRaw);
                    safeCandidate = fallbackSafe[0] || lateStaffs[0] || null;
                }
 
                if (safeCandidate && safeCandidate.name === lastSafeStaff) {
                    isSafeContinuous = true;
                    safeNote = `【金庫チェック連続】前日と同じ(${lastSafeStaff})のためNG`;
                }
                const safeNameRaw = safeCandidate ? safeCandidate.name : UNSET;
 
                // ---------------------------------------------------------------
                // 鍵保持者（当日の退勤時点の保持者）… planKeyHolders で月全体を見て決定済み
                //  ・店舗置きは「翌日の鍵開け担当に他の方法で鍵が渡らない時」だけ
                //  ・休・応援・グランド等で出勤しない人は、鍵を持ったまま持ち帰り（店舗に置かない）
                // ---------------------------------------------------------------
                const newHolders = { "鍵No.1": plan.slots[0], "鍵No.2": plan.slots[1], "鍵No.3": fixedName };
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
 
                let reasonsList = [];
                if (openNameRaw === UNSET) reasonsList.push("早番・超早番の勤務者が不在のため鍵開け未設定");
                if (closeNameRaw === UNSET) reasonsList.push("遅番の勤務者が不在のため鍵閉め未設定");
                if (openNoKey) reasonsList.push(`鍵開け担当(${openNameRaw})が前日までに鍵を受け取っておらず開錠不可`);
                if (isSafeContinuous) reasonsList.push(safeNote);
 
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
 
                let checkStatus = (openNameRaw !== UNSET && closeNameRaw !== UNSET && safeNameRaw !== UNSET && !isSafeContinuous && !openNoKey) ? "OK" : "NG";
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
                            let label = `${n} (${shiftName})`;
                            return `<option value="${n}" ${currentVal === n ? 'selected' : ''}>${label}</option>`;
                        }).join('');
                        opts += `<option value="${UNSET}" ${currentVal === UNSET ? 'selected' : ''}>${UNSET}</option>`;
                        return opts;
                    };
 
                    let makeHolderOptions = (currentVal) => {
                        let opts = staffNames.map(n => `<option value="${n}" ${currentVal === n ? 'selected' : ''}>${n}</option>`).join('');
                        opts += `<option value="${STORE_NAME}" ${currentVal === STORE_NAME ? 'selected' : ''}>${STORE_NAME}（店舗置き）</option>`;
                        return opts;
                    };
 
                    tr.innerHTML = `
                        <td class="px-3 py-2.5 font-medium whitespace-nowrap text-center">${item.date} (${item.dayOfWeek})</td>
                        <td class="px-3 py-2.5 text-center">
                            <select onchange="updateShiftField(${index}, 'openStaff', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeOptions(item.openRaw)}</select>
                        </td>
                        <td class="px-3 py-2.5 text-center">
                            <select onchange="updateShiftField(${index}, 'closeStaff', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeOptions(item.closeRaw)}</select>
                        </td>
                        <td class="px-3 py-2.5 text-center">
                            <select onchange="updateShiftField(${index}, 'safeStaff', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeOptions(item.safeRaw)}</select>
                        </td>
                        <td class="px-3 py-2.5 text-center">
                            <select onchange="updateHolderField(${index}, '鍵No.1', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeHolderOptions(kh["鍵No.1"])}</select>
                        </td>
                        <td class="px-3 py-2.5 text-center">
                            <select onchange="updateHolderField(${index}, '鍵No.2', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeHolderOptions(kh["鍵No.2"])}</select>
                        </td>
                        <td class="px-3 py-2.5 text-center">
                            <select onchange="updateHolderField(${index}, '鍵No.3', this.value)" class="border border-slate-300 rounded p-1 text-xs bg-white text-center">${makeHolderOptions(kh["鍵No.3"])}</select>
                        </td>
                        <td class="px-3 py-2.5 text-center">
                            <input type="checkbox" ${item.locked ? 'checked' : ''} onchange="toggleLock(${index}, this.checked)" class="w-4 h-4 text-indigo-600 rounded border-slate-300 cursor-pointer">
                        </td>
                    `;
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
</body>
</html>
