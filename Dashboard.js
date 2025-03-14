/**
 * マインドエンジニアリング・コーチング管理システム
 * ダッシュボード・ウェブアプリケーション
 * 
 * システムの各機能を統合したウェブUIを提供します。
 * サーバーサイドレンダリング方式を採用しています。
 */

/**
 * ウェブアプリとしてドプロイしたときに、GETリクエストを処理する関数
 * ページ読み込み時にサーバー側でデータを取得し、HTMLに埋め込みます
 * @return {HtmlOutput} HTML出力
 */
function doGet() {
  // ダッシュボードに必要なデータを取得（実データを使用）
  const dashboardData = getDashboardData();
  
  // データをJSON形式に変換
  const dataJSON = JSON.stringify(dashboardData);
  
  // ダッシュボードHTMLテンプレートを取得
  let dashboardHtml = createDashboardHtml(dataJSON);
  
  // HTMLを出力
  return HtmlService.createHtmlOutput(dashboardHtml)
    .setTitle('MEC管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ダッシュボード用のHTMLを生成する
 * @param {string} initialData - JSONで直列化されたダッシュボードデータ
 * @return {string} HTML文字列
 */
function createDashboardHtml(initialData) {
  // ダッシュボードのHTML内容
  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>MEC管理システム - ダッシュボード</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    :root {
      --primary-color: #c50502;
      --primary-dark: #9c0401;
      --primary-light: #e64644;
      --accent-color: #333333;
      --light-bg: #f8f9fa;
      --border-color: #dee2e6;
    }
    
    body {
      font-family: 'Noto Sans JP', sans-serif;
      background-color: var(--light-bg);
    }
    
    .navbar-brand {
      font-weight: bold;
      letter-spacing: 1px;
    }
    
    .bg-primary {
      background-color: var(--primary-color) !important;
    }
    
    .btn-primary {
      background-color: var(--primary-color);
      border-color: var(--primary-color);
    }
    
    .btn-primary:hover {
      background-color: var(--primary-dark);
      border-color: var(--primary-dark);
    }
    
    .btn-outline-primary {
      color: var(--primary-color);
      border-color: var(--primary-color);
    }
    
    .btn-outline-primary:hover {
      background-color: var(--primary-color);
      border-color: var(--primary-color);
    }
    
    .card {
      border-radius: 10px;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
      transition: transform 0.3s;
      overflow: hidden;
      margin-bottom: 20px;
    }
    
    .card:hover {
      transform: translateY(-5px);
    }
    
    .card-header {
      background-color: var(--primary-color);
      color: white;
      font-weight: bold;
    }
    
    .stat-card {
      border-left: 4px solid var(--primary-color);
      text-align: center;
      padding: 20px;
      height: 100%;
    }
    
    .stat-card i {
      font-size: 32px;
      color: var(--primary-color);
      margin-bottom: 10px;
    }
    
    .stat-card .stat-value {
      font-size: 28px;
      font-weight: bold;
    }
    
    .stat-card .stat-label {
      color: #666;
      font-size: 14px;
    }
    
    .sidebar {
      height: 100vh;
      background-color: white;
      border-right: 1px solid var(--border-color);
      position: sticky;
      top: 0;
    }
    
    .sidebar-link {
      padding: 10px 20px;
      display: flex;
      align-items: center;
      color: var(--accent-color);
      text-decoration: none;
      transition: all 0.3s;
    }
    
    .sidebar-link i {
      margin-right: 10px;
      width: 20px;
      text-align: center;
    }
    
    .sidebar-link:hover, .sidebar-link.active {
      background-color: var(--light-bg);
      color: var(--primary-color);
      border-left: 4px solid var(--primary-color);
    }
    
    .content-wrapper {
      padding: 20px;
    }
    
    .section-title {
      font-size: 24px;
      font-weight: bold;
      margin-bottom: 20px;
      color: var(--accent-color);
      border-bottom: 2px solid var(--primary-color);
      display: inline-block;
      padding-bottom: 5px;
    }
    
    .session-item {
      padding: 10px 15px;
      border-left: 3px solid var(--primary-color);
      margin-bottom: 10px;
      background-color: white;
      border-radius: 5px;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }
    
    .session-time {
      font-weight: bold;
      color: var(--primary-color);
    }
    
    .client-status {
      display: inline-block;
      padding: 3px 10px;
      border-radius: 20px;
      font-size: 12px;
      color: white;
      background-color: #6c757d;
    }
    
    .client-status.inquiry {
      background-color: #17a2b8;
    }
    
    .client-status.trial {
      background-color: #ffc107;
      color: #333;
    }
    
    .client-status.contract {
      background-color: #28a745;
    }
    
    .client-status.completed {
      background-color: #6c757d;
    }
    
    .chart-container {
      height: 250px;
    }
    
    .task-item {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 15px;
      border-bottom: 1px solid var(--border-color);
    }
    
    .task-item:last-child {
      border-bottom: none;
    }
    
    .task-item .form-check-input:checked + .form-check-label {
      text-decoration: line-through;
      color: #6c757d;
    }
    
    .task-date {
      font-size: 12px;
      color: #6c757d;
    }
    
    .calendar-event {
      padding: 5px 10px;
      margin-bottom: 5px;
      border-radius: 5px;
      background-color: var(--primary-light);
      color: white;
      font-size: 12px;
    }
    
    .calendar-day {
      text-align: center;
      padding: 5px;
      border: 1px solid var(--border-color);
      height: 100px;
      background-color: white;
    }
    
    .calendar-day.today {
      background-color: #e9f5ff;
    }
    
    .calendar-day-number {
      font-weight: bold;
      margin-bottom: 5px;
    }
    
    .sidebar-header {
      padding: 15px;
      background-color: var(--primary-color);
      color: white;
      font-weight: bold;
      text-align: center;
    }
    
    .user-info {
      padding: 15px;
      border-bottom: 1px solid var(--border-color);
      text-align: center;
    }
    
    .user-name {
      font-weight: bold;
      margin-top: 10px;
    }
    
    .user-role {
      font-size: 12px;
      color: #6c757d;
    }
    
    .app-footer {
      font-size: 12px;
      color: #6c757d;
      text-align: center;
      padding: 20px 0;
      margin-top: 40px;
      border-top: 1px solid var(--border-color);
    }
    
    .loading {
      display: flex;
      align-items: center;
      justify-content: center;
      height: 100px;
    }
    
    .spinner-border {
      color: var(--primary-color);
    }
    
    @media (max-width: 768px) {
      .sidebar {
        position: fixed;
        z-index: 1030;
        left: -100%;
        width: 100%;
        transition: all 0.3s;
      }
      
      .sidebar.show {
        left: 0;
      }
      
      .content-wrapper {
        margin-left: 0;
      }
      
      .navbar-brand {
        margin-left: 40px;
      }
    }
  </style>
</head>
<body>
  <div class="container-fluid">
    <div class="row">
      <!-- サイドバー -->
      <div class="col-md-2 col-lg-2 d-none d-md-block p-0 sidebar">
        <div class="sidebar-header">
          <h5 class="mb-0">MEC管理システム</h5>
        </div>
        <div class="user-info">
          <div class="user-avatar">
            <img src="https://via.placeholder.com/50" alt="User Avatar" class="rounded-circle">
          </div>
          <div class="user-name">森山雄太</div>
          <div class="user-role">管理者</div>
        </div>
        <div class="sidebar-menu mt-3">
          <a href="#" class="sidebar-link active" data-page="dashboard">
            <i class="fas fa-tachometer-alt"></i>
            ダッシュボード
          </a>
          <a href="#" class="sidebar-link" data-page="clients">
            <i class="fas fa-users"></i>
            クライアント管理
          </a>
          <a href="#" class="sidebar-link" data-page="sessions">
            <i class="fas fa-calendar-alt"></i>
            セッション管理
          </a>
          <a href="#" class="sidebar-link" data-page="payments">
            <i class="fas fa-yen-sign"></i>
            支払い管理
          </a>
          <a href="#" class="sidebar-link" data-page="emails">
            <i class="fas fa-envelope"></i>
            メール管理
          </a>
          <a href="#" class="sidebar-link" data-page="documents">
            <i class="fas fa-file-contract"></i>
            契約書管理
          </a>
          <a href="#" class="sidebar-link" data-page="reports">
            <i class="fas fa-chart-bar"></i>
            レポート
          </a>
          <a href="#" class="sidebar-link" data-page="settings">
            <i class="fas fa-cog"></i>
            設定
          </a>
        </div>
      </div>
      
      <!-- メインコンテンツ -->
      <div class="col-md-10 col-lg-10 ms-sm-auto p-0">
        <!-- ナビゲーションバー -->
        <nav class="navbar navbar-expand-lg navbar-dark bg-primary">
          <div class="container-fluid">
            <button class="navbar-toggler d-md-none" type="button" id="sidebarToggle">
              <span class="navbar-toggler-icon"></span>
            </button>
            <a class="navbar-brand d-md-none" href="#">MEC管理システム</a>
            <div class="d-flex">
              <div class="me-3">
                <a href="#" class="text-white"><i class="fas fa-bell"></i></a>
              </div>
              <div class="dropdown">
                <a href="#" class="text-white dropdown-toggle" id="dropdownMenuButton" data-bs-toggle="dropdown">
                  <i class="fas fa-user-circle"></i>
                </a>
                <ul class="dropdown-menu dropdown-menu-end" aria-labelledby="dropdownMenuButton">
                  <li><a class="dropdown-item" href="#">プロフィール</a></li>
                  <li><a class="dropdown-item" href="#">設定</a></li>
                  <li><hr class="dropdown-divider"></li>
                  <li><a class="dropdown-item" href="#">ログアウト</a></li>
                </ul>
              </div>
            </div>
          </div>
        </nav>
        
        <!-- コンテンツ -->
        <div class="content-wrapper">
          <!-- ページコンテンツはここに動的に追加されます -->
          <div id="page-content">
            <!-- ダッシュボード -->
            <div id="dashboard-page">
              <div class="d-flex justify-content-between align-items-center mb-4">
                <h1 class="h2">ダッシュボード</h1>
                <div>
                  <span class="text-muted me-2">最終更新: <span id="last-update-time">読み込み中...</span></span>
                  <button class="btn btn-sm btn-outline-primary" id="refresh-button">
                    <i class="fas fa-sync-alt"></i> 更新
                  </button>
                </div>
              </div>
              
              <!-- 統計カード -->
              <div class="row mb-4">
                <div class="col-md-3 mb-3">
                  <div class="card h-100">
                    <div class="card-body p-0">
                      <div class="stat-card">
                        <i class="fas fa-users"></i>
                        <div class="stat-value" id="active-clients-count">-</div>
                        <div class="stat-label">アクティブクライアント</div>
                      </div>
                    </div>
                  </div>
                </div>
                <div class="col-md-3 mb-3">
                  <div class="card h-100">
                    <div class="card-body p-0">
                      <div class="stat-card">
                        <i class="fas fa-calendar-check"></i>
                        <div class="stat-value" id="today-sessions-count">-</div>
                        <div class="stat-label">今日のセッション</div>
                      </div>
                    </div>
                  </div>
                </div>
                <div class="col-md-3 mb-3">
                  <div class="card h-100">
                    <div class="card-body p-0">
                      <div class="stat-card">
                        <i class="fas fa-yen-sign"></i>
                        <div class="stat-value" id="monthly-sales">-</div>
                        <div class="stat-label">今月の売上</div>
                      </div>
                    </div>
                  </div>
                </div>
                <div class="col-md-3 mb-3">
                  <div class="card h-100">
                    <div class="card-body p-0">
                      <div class="stat-card">
                        <i class="fas fa-tasks"></i>
                        <div class="stat-value" id="pending-tasks-count">-</div>
                        <div class="stat-label">未完了タスク</div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              
              <!-- コンテンツエリア -->
              <div class="row">
                <!-- 左カラム -->
                <div class="col-md-8">
                  <!-- 今日のセッション -->
                  <div class="card mb-4">
                    <div class="card-header d-flex justify-content-between align-items-center">
                      <h5 class="mb-0">今日のセッション</h5>
                      <a href="#" class="btn btn-sm btn-outline-light" data-page="sessions">すべて表示</a>
                    </div>
                    <div class="card-body" id="today-sessions-container">
                      <div class="loading">
                        <div class="spinner-border" role="status">
                          <span class="visually-hidden">Loading...</span>
                        </div>
                      </div>
                    </div>
                  </div>
                  
                  <!-- 売上グラフ -->
                  <div class="card mb-4">
                    <div class="card-header">
                      <h5 class="mb-0">売上推移</h5>
                    </div>
                    <div class="card-body">
                      <div class="chart-container">
                        <canvas id="salesChart"></canvas>
                      </div>
                    </div>
                  </div>
                  
                  <!-- 最近のクライアント -->
                  <div class="card">
                    <div class="card-header d-flex justify-content-between align-items-center">
                      <h5 class="mb-0">最近のクライアント</h5>
                      <a href="#" class="btn btn-sm btn-outline-light" data-page="clients">すべて表示</a>
                    </div>
                    <div class="card-body" id="recent-clients-container">
                      <div class="loading">
                        <div class="spinner-border" role="status">
                          <span class="visually-hidden">Loading...</span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
                
                <!-- 右カラム -->
                <div class="col-md-4">
                  <!-- タスク -->
                  <div class="card mb-4">
                    <div class="card-header d-flex justify-content-between align-items-center">
                      <h5 class="mb-0">タスク</h5>
                      <button class="btn btn-sm btn-outline-light">
                        <i class="fas fa-plus"></i> 追加
                      </button>
                    </div>
                    <div class="card-body" id="tasks-container">
                      <div class="loading">
                        <div class="spinner-border" role="status">
                          <span class="visually-hidden">Loading...</span>
                        </div>
                      </div>
                    </div>
                    <div class="card-footer text-center">
                      <a href="#" class="btn btn-sm btn-outline-secondary">すべてのタスクを表示</a>
                    </div>
                  </div>
                  
                  <!-- クライアント状況 -->
                  <div class="card mb-4">
                    <div class="card-header">
                      <h5 class="mb-0">クライアント状況</h5>
                    </div>
                    <div class="card-body">
                      <div class="chart-container">
                        <canvas id="clientStatusChart"></canvas>
                      </div>
                    </div>
                  </div>
                  
                  <!-- カレンダー（週表示） -->
                  <div class="card">
                    <div class="card-header d-flex justify-content-between align-items-center">
                      <h5 class="mb-0">今週のスケジュール</h5>
                      <a href="#" class="btn btn-sm btn-outline-light" data-page="sessions">月表示</a>
                    </div>
                    <div class="card-body" id="weekly-calendar-container">
                      <div class="loading">
                        <div class="spinner-border" role="status">
                          <span class="visually-hidden">Loading...</span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              
              <!-- フッター -->
              <footer class="app-footer">
                <p>© 2025 マインドエンジニアリング・コーチング 管理システム</p>
              </footer>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
  
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.bundle.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  
  <script>
    // サーバーから取得した初期データ
    const initialData = ${initialData};
    
    // ユーティリティ関数
    function formatCurrency(amount) {
      return '¥' + amount.toLocaleString();
    }
    
    function formatDateTime(dateTime) {
      const d = new Date(dateTime);
      const year = d.getFullYear();
      const month = (d.getMonth() + 1).toString().padStart(2, '0');
      const day = d.getDate().toString().padStart(2, '0');
      const hours = d.getHours().toString().padStart(2, '0');
      const minutes = d.getMinutes().toString().padStart(2, '0');
      return \`\${year}-\${month}-\${day} \${hours}:\${minutes}\`;
    }
    
    function formatShortDate(date) {
      const d = new Date(date);
      const month = (d.getMonth() + 1).toString();
      const day = d.getDate().toString();
      return \`\${month}/\${day}\`;
    }
    
    // ページが読み込まれたときの処理
    document.addEventListener('DOMContentLoaded', function() {
      // サイドバートグルボタンの処理
      const sidebarToggle = document.getElementById('sidebarToggle');
      const sidebar = document.querySelector('.sidebar');
      
      if (sidebarToggle) {
        sidebarToggle.addEventListener('click', function() {
          sidebar.classList.toggle('show');
        });
      }
      
      // 更新ボタンの処理
      const refreshButton = document.getElementById('refresh-button');
      if (refreshButton) {
        refreshButton.addEventListener('click', function() {
          // サーバーから最新データを取得
          google.script.run.withSuccessHandler(handleDashboardData).getDashboardData();
        });
      }
      
      // 最終更新時間を設定
      updateLastUpdateTime();
      
      // 初期データで画面を更新
      handleDashboardData(initialData);
      
      // 売上グラフとクライアントステータスグラフを初期化
      initCharts();
      
      // 画面をデータで更新
      updateCharts(initialData.salesData, initialData.clientStatusData);
    });
    
    // 最終更新時間を更新
    function updateLastUpdateTime() {
      const lastUpdateTimeElement = document.getElementById('last-update-time');
      if (lastUpdateTimeElement) {
        const now = new Date();
        lastUpdateTimeElement.textContent = formatDateTime(now);
      }
    }
    
    // ダッシュボードデータを処理
    function handleDashboardData(data) {
      // 基本統計を更新
      document.getElementById('active-clients-count').textContent = data.activeClientsCount;
      document.getElementById('today-sessions-count').textContent = data.todaySessionsCount;
      document.getElementById('monthly-sales').textContent = formatCurrency(data.monthlySales);
      document.getElementById('pending-tasks-count').textContent = data.pendingTasksCount;
      
      // 今日のセッションを更新
      updateTodaySessions(data.todaySessions);
      
      // 最近のクライアントを更新
      updateRecentClients(data.recentClients);
      
      // タスクを更新
      updateTasks(data.tasks);
      
      // 週間カレンダーを更新
      updateWeeklyCalendar(data.weeklyCalendar);
      
      // グラフを更新
      updateCharts(data.salesData, data.clientStatusData);
      
      // 最終更新時間を更新
      updateLastUpdateTime();
    }
    
    // 今日のセッションを更新
    function updateTodaySessions(sessions) {
      const container = document.getElementById('today-sessions-container');
      if (!container) return;
      
      container.innerHTML = '';
      
      if (sessions.length === 0) {
        container.innerHTML = '<p class="text-center">今日のセッションはありません。</p>';
        return;
      }
      
      sessions.forEach(session => {
        let statusClass = '';
        if (session.status === 'トライアル前') statusClass = 'trial';
        else if (session.status === '契約中') statusClass = 'contract';
        else if (session.status === '問い合わせ') statusClass = 'inquiry';
        
        const sessionHtml = \`
          <div class="session-item">
            <div class="d-flex justify-content-between align-items-center">
              <div>
                <span class="session-time">\${session.time}</span>
                <h6 class="mb-0">\${session.clientName} 様</h6>
                <span class="client-status \${statusClass}">\${session.status}</span>
              </div>
              <div>
                <span class="badge bg-\${session.sessionType === 'オンライン' ? 'primary' : 'success'}">\${session.sessionType}</span>
                \${session.meetUrl ? \`<a href="\${session.meetUrl}" target="_blank" class="btn btn-sm btn-outline-primary ms-2">
                  <i class="fas fa-video"></i> Google Meet
                </a>\` : ''}
              </div>
            </div>
          </div>
        \`;
        
        container.innerHTML += sessionHtml;
      });
    }
    
    // 最近のクライアントを更新
    function updateRecentClients(clients) {
      const container = document.getElementById('recent-clients-container');
      if (!container) return;
      
      container.innerHTML = '';
      
      if (clients.length === 0) {
        container.innerHTML = '<p class="text-center">クライアント情報がありません。</p>';
        return;
      }
      
      let tableHtml = \`
        <div class="table-responsive">
          <table class="table table-hover">
            <thead>
              <tr>
                <th>名前</th>
                <th>ステータス</th>
                <th>登録日</th>
                <th>セッション形式</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
      \`;
      
      clients.forEach(client => {
        let statusClass = '';
        if (client.status === 'トライアル前' || client.status === 'トライアル済') statusClass = 'trial';
        else if (client.status === '契約中') statusClass = 'contract';
        else if (client.status === '問い合わせ') statusClass = 'inquiry';
        
        tableHtml += \`
          <tr>
            <td>\${client.name}</td>
            <td><span class="client-status \${statusClass}">\${client.status}</span></td>
            <td>\${client.registrationDate}</td>
            <td>\${client.sessionType}</td>
            <td><a href="#" class="btn btn-sm btn-outline-primary">詳細</a></td>
          </tr>
        \`;
      });
      
      tableHtml += \`
            </tbody>
          </table>
        </div>
      \`;
      
      container.innerHTML = tableHtml;
    }
    
    // タスクを更新
    function updateTasks(tasks) {
      const container = document.getElementById('tasks-container');
      if (!container) return;
      
      container.innerHTML = '';
      
      if (tasks.length === 0) {
        container.innerHTML = '<p class="text-center">タスクはありません。</p>';
        return;
      }
      
      tasks.forEach(task => {
        const taskHtml = \`
          <div class="task-item">
            <div class="form-check">
              <input class="form-check-input" type="checkbox" id="task\${task.id}" \${task.completed ? 'checked' : ''}>
              <label class="form-check-label" for="task\${task.id}">
                \${task.description}
              </label>
            </div>
            <div class="task-date">\${task.dueDate}</div>
          </div>
        \`;
        
        container.innerHTML += taskHtml;
      });
      
      // チェックボックスの変更イベントを追加
      const checkboxes = container.querySelectorAll('input[type="checkbox"]');
      checkboxes.forEach(checkbox => {
        checkbox.addEventListener('change', function() {
          const taskId = parseInt(this.id.replace('task', ''));
          // GAS関数を呼び出して、タスク完了状態を更新
          // 例: google.script.run.updateTaskStatus(taskId, this.checked);
          console.log(\`タスク \${taskId} の状態を \${this.checked ? '完了' : '未完了'} に更新\`);
        });
      });
    }
    
    // 週間カレンダーを更新
    function updateWeeklyCalendar(calendarData) {
      const container = document.getElementById('weekly-calendar-container');
      if (!container) return;
      
      container.innerHTML = '';
      
      // 曜日の列を作成
      let calendarHtml = \`
        <div class="row text-center fw-bold mb-2">
          <div class="col">月</div>
          <div class="col">火</div>
          <div class="col">水</div>
          <div class="col">木</div>
          <div class="col">金</div>
        </div>
        <div class="row mb-2">
      \`;
      
      // 各日の列を作成
      calendarData.days.forEach(day => {
        calendarHtml += \`
          <div class="col calendar-day \${day.isToday ? 'today' : ''}">
            <div class="calendar-day-number">\${day.date}</div>
        \`;
        
        // イベントを追加
        day.events.forEach(event => {
          calendarHtml += \`
            <div class="calendar-event">\${event.time} \${event.clientName}</div>
          \`;
        });
        
        calendarHtml += \`
          </div>
        \`;
      });
      
      calendarHtml += \`
        </div>
        <div class="text-center mt-3">
          <a href="#" class="btn btn-sm btn-primary" data-page="sessions">
            <i class="fas fa-calendar-alt"></i> カレンダーを表示
          </a>
        </div>
      \`;
      
      container.innerHTML = calendarHtml;
    }
    
    // グラフを初期化
    function initCharts() {
      // 売上グラフ
      const salesCtx = document.getElementById('salesChart').getContext('2d');
      window.salesChart = new Chart(salesCtx, {
        type: 'line',
        data: {
          labels: [],
          datasets: [{
            label: '売上',
            data: [],
            backgroundColor: 'rgba(197, 5, 2, 0.2)',
            borderColor: '#c50502',
            borderWidth: 2,
            tension: 0.3
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: {
              beginAtZero: true,
              ticks: {
                callback: function(value) {
                  return '¥' + value.toLocaleString();
                }
              }
            }
          }
        }
      });
      
      // クライアント状況グラフ
      const clientCtx = document.getElementById('clientStatusChart').getContext('2d');
      window.clientStatusChart = new Chart(clientCtx, {
        type: 'doughnut',
        data: {
          labels: [],
          datasets: [{
            data: [],
            backgroundColor: [
              '#17a2b8',
              '#ffc107',
              '#fd7e14',
              '#28a745',
              '#6c757d',
              '#dc3545'
            ],
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'bottom'
            }
          },
          cutout: '70%'
        }
      });
    }
    
    // グラフを更新
    function updateCharts(salesData, clientStatusData) {
      // 売上グラフを更新
      if (window.salesChart) {
        window.salesChart.data.labels = salesData.labels;
        window.salesChart.data.datasets[0].data = salesData.data;
        window.salesChart.update();
      }
      
      // クライアント状況グラフを更新
      if (window.clientStatusChart) {
        window.clientStatusChart.data.labels = clientStatusData.labels;
        window.clientStatusChart.data.datasets[0].data = clientStatusData.data;
        window.clientStatusChart.update();
      }
    }
  </script>
</body>
</html>`;
}

/**
 * ダッシュボード用のHTMLファイルを作成する
 * GASのコードエディタからこの関数を実行することで、dashboard.htmlファイルを自動作成します。
 */
function createDashboardHtmlFile() {
  // ダッシュボードのHTML内容
  const dashboardHtml = createDashboardHtml("{}");
  
  // GASプロジェクトにHTMLファイルを作成
  try {
    const htmlFile = DriveApp.getFileById(ScriptApp.getScriptId())
      .getParent()
      .createFile('dashboard.html', dashboardHtml, MimeType.HTML);
    
    return {
      success: true,
      message: 'dashboard.htmlファイルを作成しました。',
      fileId: htmlFile.getId(),
      fileUrl: htmlFile.getUrl()
    };
  } catch (error) {
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.message
    };
  }
}

/**
 * ダッシュボード用のデータを取得
 * 実際のデータベースからデータを取得します
 * @return {Object} ダッシュボードのデータ
 */
function getDashboardData() {
  // クライアントデータを取得
  const activeClients = ClientManager.getAllClients(true);
  const activeClientsCount = activeClients.length;
  
  // 今日のセッションを取得
  const todaySessions = SessionManager.getTodaySessions();
  const todaySessionsCount = todaySessions.length;
  
  // 今月の売上を取得
  const now = new Date();
  const monthlySales = PaymentManager.getMonthlySales(now.getFullYear(), now.getMonth() + 1);
  
  // クライアントステータスの集計
  const clientStatusSummary = ClientManager.getClientStatusSummary();
  
  // クライアントステータスグラフデータを作成
  const clientStatusData = {
    labels: ['問い合わせ', 'トライアル前', 'トライアル済', '契約中', '完了', '中断'],
    data: [
      clientStatusSummary['問い合わせ'] || 0,
      clientStatusSummary['トライアル前'] || 0,
      clientStatusSummary['トライアル済'] || 0,
      clientStatusSummary['契約中'] || 0,
      clientStatusSummary['完了'] || 0,
      clientStatusSummary['中断'] || 0
    ]
  };
  
  // 売上データを取得
  const yearData = PaymentManager.getYearlySalesData();
  const salesData = {
    labels: yearData.slice(-6).map(monthData => monthData.monthName),
    data: yearData.slice(-6).map(monthData => monthData.sales)
  };
  
  // 今日のセッション情報を整形
  const formattedTodaySessions = todaySessions.map(session => {
    // セッション時間を計算
    const startTime = new Date(session.予定日時);
    const endTime = new Date(startTime);
    endTime.setMinutes(endTime.getMinutes() + parseInt(Utilities.getSetting('SESSION_DURATION', '30')));
    
    // 時間の文字列を作成
    const timeStr = `${startTime.getHours().toString().padStart(2, '0')}:${startTime.getMinutes().toString().padStart(2, '0')} - ${endTime.getHours().toString().padStart(2, '0')}:${endTime.getMinutes().toString().padStart(2, '0')}`;
    
    // クライアント情報を取得
    const client = ClientManager.findClientById(session.クライアントID);
    
    return {
      time: timeStr,
      clientName: client ? client.お名前 : '不明なクライアント',
      status: client ? client.ステータス : '不明',
      sessionType: client ? client.希望セッション形式 : '不明',
      meetUrl: session['Google Meet URL'] || ''
    };
  });
  
  // 最近のクライアントを取得
  // すべてのクライアントを取得して、登録日でソート
  const allClients = ClientManager.getAllClients();
  allClients.sort((a, b) => new Date(b.タイムスタンプ) - new Date(a.タイムスタンプ));
  
  // 最新の4件を使用
  const recentClients = allClients.slice(0, 4).map(client => {
    const regDate = new Date(client.タイムスタンプ);
    return {
      name: client.お名前,
      status: client.ステータス,
      registrationDate: `${regDate.getFullYear()}/${(regDate.getMonth() + 1).toString().padStart(2, '0')}/${regDate.getDate().toString().padStart(2, '0')}`,
      sessionType: client.希望セッション形式 || '未定'
    };
  });
  
  // タスク機能は未実装のため、サンプルデータを生成
  const tasks = [
    {
      id: 1,
      description: 'クライアントへのリマインダーメールを送信',
      dueDate: 'Today',
      completed: false
    },
    {
      id: 2,
      description: 'セッション記録の入力',
      dueDate: 'Today',
      completed: false
    },
    {
      id: 3,
      description: '契約書のテンプレート作成',
      dueDate: 'Tomorrow',
      completed: false
    },
    {
      id: 4,
      description: '月次レポートの準備',
      dueDate: '3/20',
      completed: false
    }
  ];
  
  // 週間カレンダーを作成
  const today = new Date();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - today.getDay() + 1); // 月曜日に設定
  
  const weeklyCalendar = {
    days: []
  };
  
  // 月曜日から金曜日までのデータを作成
  for (let i = 0; i < 5; i++) {
    const currentDay = new Date(startOfWeek);
    currentDay.setDate(startOfWeek.getDate() + i);
    
    // この日のセッションを取得
    const daySessions = SessionManager.getSessionsByDate(currentDay);
    
    const dayEvents = daySessions.map(session => {
      const client = ClientManager.findClientById(session.クライアントID);
      const sessionTime = new Date(session.予定日時);
      return {
        time: `${sessionTime.getHours().toString().padStart(2, '0')}:${sessionTime.getMinutes().toString().padStart(2, '0')}`,
        clientName: client ? `${client.お名前}様` : '不明なクライアント'
      };
    });
    
    weeklyCalendar.days.push({
      date: `${currentDay.getMonth() + 1}/${currentDay.getDate()}`,
      isToday: currentDay.getDate() === today.getDate() && 
               currentDay.getMonth() === today.getMonth() && 
               currentDay.getFullYear() === today.getFullYear(),
      events: dayEvents
    });
  }
  
  // ダッシュボードデータをまとめて返す
  return {
    activeClientsCount,
    todaySessionsCount,
    monthlySales,
    pendingTasksCount: tasks.filter(t => !t.completed).length,
    todaySessions: formattedTodaySessions,
    recentClients,
    tasks,
    weeklyCalendar,
    salesData,
    clientStatusData
  };
}