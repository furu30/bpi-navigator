/**
 * 業務プロセス改善ナビゲーター - メインアプリケーション
 * ローカルストレージベースのSPAアプリケーション
 */
;(function() {
  'use strict';

  // ==========================================
  // データストア
  // ==========================================
  const STORAGE_KEY = 'bpiNavi_data';

  const DEFAULT_PROCESS_CATEGORIES = [
    '受注', '生産管理', '製造', '品質管理', '出荷', '購買', '設計開発', '総務', '経理', 'その他'
  ];

  const PROBLEM_CATEGORIES = [
    { id: 'duplicate', name: '転記ミス・二重入力', desc: '同じ情報を複数箇所に入力する無駄' },
    { id: 'waiting', name: '待ち時間・遅延', desc: '前工程や他者の対応を待つ時間' },
    { id: 'nostandard', name: '標準化不足', desc: '手順やルールが統一されていない' },
    { id: 'paper', name: '紙・手作業', desc: 'デジタル化されていない作業' },
    { id: 'personal', name: '属人化', desc: '特定の人しかできない作業' },
    { id: 'search', name: '検索・探索', desc: '必要な物や情報を探す時間' },
    { id: 'communication', name: 'コミュニケーション', desc: '情報伝達・調整に関する問題' },
    { id: 'system', name: 'システム連携不足', desc: 'システム間のデータ連携がない' }
  ];

  const IMPROVEMENT_TEMPLATES = {
    duplicate: ['RPAによる自動転記', 'マスタデータの一元管理', 'システム連携（API）の導入'],
    waiting: ['工程の並列化', '承認プロセスの簡略化', '自動通知の仕組み導入'],
    nostandard: ['作業手順書の作成', 'チェックリストの導入', '定例レビューの実施'],
    paper: ['デジタルフォームの導入', 'タブレット入力への移行', 'OCRによる自動読み取り'],
    personal: ['マニュアル・動画の整備', 'クロストレーニングの実施', 'ナレッジベースの構築'],
    search: ['5S活動の実施', 'ラベル・色分けルール整備', '検索システムの導入'],
    communication: ['定例ミーティングの設定', 'チャットツールの活用', '情報共有ボードの設置'],
    system: ['データ連携ツールの導入', 'APIによるシステム統合', '共通データベースの構築']
  };

  const ECRS_DEFINITIONS = [
    {
      key: 'E', name: '排除（Eliminate）', color: 'e',
      questions: [
        'この作業をなくしたら何が困りますか？',
        'この作業は最終成果物に貢献していますか？'
      ],
      criteria: '「困らない」→ 排除候補'
    },
    {
      key: 'C', name: '結合（Combine）', color: 'c',
      questions: [
        'この作業と一緒にできる作業はありますか？',
        '同じ担当者が連続して行える作業はありますか？'
      ],
      criteria: '同一担当・同一場所 → 結合候補'
    },
    {
      key: 'R', name: '交換（Rearrange）', color: 'r',
      questions: [
        'この作業の順番を変えたら効率は上がりますか？',
        '他の人がやったほうが効率的ですか？'
      ],
      criteria: 'ボトルネック前後 → 交換候補'
    },
    {
      key: 'S', name: '簡素化（Simplify）', color: 's',
      questions: [
        'この作業をもっと簡単にする方法はありますか？',
        'ツールで自動化できる部分はありますか？'
      ],
      criteria: '手作業・紙作業 → 簡素化候補'
    }
  ];

  /** @returns {object} default data structure */
  function createDefaultData() {
    return {
      projects: [],
      companies: {},
      processCategories: [...DEFAULT_PROCESS_CATEGORIES],
      currentProjectId: null
    };
  }

  /** 既存データにcompaniesが無い場合、プロジェクトのcompanyフィールドから自動生成 */
  function migrateData(data) {
    if (!data.companies) {
      data.companies = {};
      (data.projects || []).forEach(p => {
        if (p.company && !data.companies[p.company]) {
          data.companies[p.company] = { description: '' };
        }
      });
    }
    return data;
  }

  // デモプレビュー状態管理（saveDataより先に宣言が必要）
  let savedDataBeforeDemo = null;
  let isDemoPreviewMode = false;

  function loadData() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return migrateData(JSON.parse(raw));
    } catch (e) { console.error('Load error:', e); }
    return createDefaultData();
  }

  function saveData(data) {
    // デモプレビュー中はlocalStorageに保存しない
    if (isDemoPreviewMode) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    } catch (e) { console.error('Save error:', e); alert('データの保存に失敗しました。'); }
  }

  let appData = loadData();

  // ==========================================
  // ユーティリティ
  // ==========================================
  function genId() { return Date.now().toString(36) + Math.random().toString(36).substr(2, 5); }

  function $(sel) { return document.querySelector(sel); }
  function $$(sel) { return document.querySelectorAll(sel); }

  function showToast(msg, type = '') {
    const t = $('#toast');
    t.textContent = msg;
    t.className = 'toast show ' + type;
    setTimeout(() => t.className = 'toast', 2500);
  }

  function openModal(id) { $('#' + id).classList.add('active'); }
  function closeModal(id) { $('#' + id).classList.remove('active'); }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>"']/g, m => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[m]));
  }

  function getCurrentProject() {
    if (!appData.currentProjectId) return null;
    return appData.projects.find(p => p.id === appData.currentProjectId) || null;
  }

  const CODE_PREFIX = {
    '受注': 'ORD', '生産管理': 'PLN', '製造': 'MFG',
    '品質管理': 'QC', '出荷': 'SHP', '購買': 'PUR',
    '設計開発': 'DEV', '総務': 'ADM', '経理': 'ACC', 'その他': 'OTH'
  };

  function generateTaskCode(category) {
    const proj = getCurrentProject();
    if (!proj) return '';
    const prefix = CODE_PREFIX[category] || 'TSK';
    const count = (proj.tasks || []).filter(t => t.category === category).length + 1;
    return `${prefix}-${String(count).padStart(3, '0')}`;
  }

  function calcTotalScore(scores) {
    if (!scores) return 0;
    return (scores.timeImpact || 0) * 0.3
      + (scores.qualityImpact || 0) * 0.25
      + (scores.frequency || 0) * 0.25
      + (scores.difficulty || 0) * 0.2;
  }

  function calcMonthlyTime(task) {
    const time = parseFloat(task.timeRequired) || 0;
    const freq = parseFloat(task.freqCount) || 1;
    const freqType = task.freqType || '月次';
    let multiplier = 1;
    if (freqType === '日次') multiplier = 20;
    else if (freqType === '週次') multiplier = 4;
    else if (freqType === '月次') multiplier = 1;
    else multiplier = 1;
    return time * freq * multiplier;
  }

  // ==========================================
  // ナビゲーション
  // ==========================================
  let currentView = 'dashboard';

  function navigate(viewName) {
    currentView = viewName;
    $$('.view').forEach(v => v.classList.remove('active'));
    const target = $(`#view-${viewName}`);
    if (target) target.classList.add('active');

    $$('.nav-item').forEach(n => {
      n.classList.remove('active');
      if (n.dataset.view === viewName) n.classList.add('active');
    });

    const titles = {
      dashboard: 'ダッシュボード',
      step1: 'Step 1: 業務プロセスの洗い出し',
      step2: 'Step 2: 問題点の可視化と分析',
      step3: 'Step 3: ECRSによる改善対象の選定',
      step4: 'Step 4: 具体的な改善策の検討',
      step5: 'Step 5: 効果検証と振り返り',
      settings: '設定・マスタ管理'
    };
    $('#topbarTitle').textContent = titles[viewName] || '';

    const proj = getCurrentProject();
    const btnExport = $('#btnExportProject');
    if (btnExport) btnExport.style.display = proj ? '' : 'none';

    if (viewName === 'dashboard') renderDashboard();
    else if (viewName === 'step1') renderStep1();
    else if (viewName === 'step2') renderStep2();
    else if (viewName === 'step3') renderStep3();
    else if (viewName === 'step4') renderStep4();
    else if (viewName === 'step5') renderStep5();
    else if (viewName === 'settings') renderSettings();

    updateNavSteps();
    updateTopbarProject();
    updateSaveProjectVisibility();

    // Close sidebar on mobile
    if (window.innerWidth <= 768) {
      $('#sidebar').classList.remove('open');
    }
  }

  function updateNavSteps() {
    const proj = getCurrentProject();
    if (!proj) return;
    const step = proj.currentStep || 1;
    $$('.nav-item').forEach(n => {
      const v = n.dataset.view;
      if (v && v.startsWith('step')) {
        const num = parseInt(v.replace('step', ''));
        n.classList.toggle('completed', num < step);
        const dot = n.querySelector('.nav-step');
        if (dot) {
          if (num < step) dot.style.background = 'var(--success)';
          else if (num === step) dot.style.background = 'var(--primary)';
          else dot.style.background = '';
        }
      }
    });
  }

  /** トップバーにプロジェクト情報とステップナビを表示/非表示 */
  function updateTopbarProject() {
    const bar = $('#topbarProjectBar');
    const stepBar = $('#stepNavBar');
    if (!bar || !stepBar) return;
    const proj = getCurrentProject();
    const isProjectView = proj && currentView !== 'dashboard' && currentView !== 'settings';

    // プロジェクト情報バー
    bar.style.display = isProjectView ? '' : 'none';
    if (isProjectView) {
      $('#topbarCompany').textContent = proj.company || '';
      $('#topbarProjectName').textContent = proj.name || '';
    }

    // ステップナビゲーションバー
    stepBar.style.display = isProjectView ? '' : 'none';
    if (!isProjectView) return;

    const stepLabels = [
      '業務洗い出し',
      '問題分析',
      'ECRS選定',
      '改善策検討',
      '効果検証'
    ];
    const curStep = proj.currentStep || 1;
    const stepsEl = $('#topbarSteps');

    // ステップボタン + 間の接続線を生成
    let html = '';
    stepLabels.forEach((label, i) => {
      const num = i + 1;
      let cls = 'step-nav-btn';
      if (num < curStep) cls += ' completed';
      else if (num === curStep) cls += ' active';
      if (currentView === 'step' + num) cls += ' current-view';

      html += `<button class="${cls}" data-step="${num}">
        <span class="step-num">${num}</span>
        <span class="step-label">${label}</span>
      </button>`;

      // 最後以外は接続線を追加
      if (i < stepLabels.length - 1) {
        const connDone = num < curStep ? ' done' : '';
        html += `<div class="step-nav-connector${connDone}"></div>`;
      }
    });
    stepsEl.innerHTML = html;

    // ステップボタンクリックでナビゲーション
    stepsEl.querySelectorAll('.step-nav-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        navigate('step' + parseInt(btn.dataset.step));
      });
    });
  }

  // ==========================================
  // ダッシュボード
  // ==========================================
  function renderDashboard() {
    const list = $('#projectList');
    const empty = $('#noProjects');

    if (appData.projects.length === 0) {
      list.style.display = 'none';
      empty.style.display = '';
      return;
    }

    empty.style.display = 'none';
    list.style.display = '';

    // 企業ごとにグループ化
    const groups = {};
    appData.projects.forEach(p => {
      const company = p.company || '（未設定）';
      if (!groups[company]) groups[company] = [];
      groups[company].push(p);
    });

    const companyKeys = Object.keys(groups);
    let html = '';
    companyKeys.forEach(company => {
      const projects = groups[company];
      // 全体ステップ平均を計算
      const avgStep = projects.reduce((s, p) => s + (p.currentStep || 1), 0) / projects.length;
      const avgProgress = Math.min(((avgStep - 1) / 5) * 100, 100);
      // 企業が1つだけなら最初から展開
      const isOpen = companyKeys.length === 1 ? ' open' : '';

      // 企業概要
      const companyInfo = (appData.companies || {})[company];
      const companyDesc = companyInfo?.description || '';
      const descHtml = companyDesc
        ? `<div class="company-description">${escapeHtml(companyDesc.length > 80 ? companyDesc.substring(0, 80) + '...' : companyDesc)}</div>`
        : '';

      html += `<div class="company-group${isOpen}">
        <div class="company-group-header" data-company="${escapeHtml(company)}">
          <div class="company-group-left">
            <span class="company-group-arrow">&#9654;</span>
            <span class="company-group-name">${escapeHtml(company)}</span>
          </div>
          <div class="company-group-right">
            <span class="company-group-count">${projects.length}件</span>
            <div class="company-group-progress">
              <div class="company-group-progress-fill" style="width:${avgProgress}%"></div>
            </div>
            <button class="btn-company-action btn-company-add-project" data-company="${escapeHtml(company)}" title="この企業にプロジェクトを追加">＋</button>
            <button class="btn-company-action btn-company-edit" data-company="${escapeHtml(company)}" title="企業情報を編集">✏️</button>
            <button class="btn-save-company" data-company="${escapeHtml(company)}" title="この企業のプロジェクトを保存">&#128190;</button>
          </div>
        </div>
        ${descHtml ? `<div class="company-header-info" data-company="${escapeHtml(company)}">${descHtml}</div>` : ''}
        <div class="company-group-body">
          <div class="project-grid">`;

      projects.forEach(p => {
        const taskCount = (p.tasks || []).length;
        const step = p.currentStep || 1;
        const progress = Math.min(((step - 1) / 5) * 100, 100);
        const steps = [1,2,3,4,5].map(s => {
          let cls = '';
          if (s < step) cls = 'completed';
          else if (s === step) cls = 'active';
          return `<div class="step-dot ${cls}">${s}</div>`;
        }).join('');

        html += `
          <div class="project-card" data-id="${p.id}">
            <div class="card-actions">
              <button class="btn-save-proj" data-save="${p.id}" title="このプロジェクトを保存">&#128190;</button>
              <button class="delete-btn" data-delete="${p.id}" title="削除">&times;</button>
            </div>
            <h3>${escapeHtml(p.name)}</h3>
            <div class="progress-bar"><div class="progress-fill" style="width:${progress}%"></div></div>
            <div style="font-size:12px;color:var(--text-secondary);">
              登録業務: ${taskCount}件 ／ 現在: Step ${step}
            </div>
            <div class="step-indicators">${steps}</div>
          </div>`;
      });

      html += `</div></div></div>`;
    });

    list.innerHTML = html;

    // アコーディオン開閉
    list.querySelectorAll('.company-group-header').forEach(header => {
      header.addEventListener('click', (e) => {
        // ボタンクリック時はアコーディオン開閉しない
        if (e.target.closest('.btn-save-company') || e.target.closest('.btn-company-action')) return;
        header.closest('.company-group').classList.toggle('open');
      });
    });

    // 企業概要行クリックでアコーディオン開閉
    list.querySelectorAll('.company-header-info').forEach(info => {
      info.addEventListener('click', () => {
        info.closest('.company-group').classList.toggle('open');
      });
    });

    // 企業編集ボタン
    list.querySelectorAll('.btn-company-edit').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        openCompanyModal(btn.dataset.company);
      });
    });

    // 企業にプロジェクト追加ボタン
    list.querySelectorAll('.btn-company-add-project').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        openNewProjectModal(btn.dataset.company);
      });
    });

    // 企業保存ボタン
    list.querySelectorAll('.btn-save-company').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const company = btn.dataset.company;
        const companyProjects = appData.projects.filter(p => p.company === company);
        const companyInfo = (appData.companies || {})[company] || {};
        const exportData = {
          _type: 'bpi_company',
          _version: 1,
          company: company,
          companyInfo: companyInfo,
          projects: companyProjects
        };
        downloadJson(exportData, `bpi_company_${sanitizeFilename(company)}_${today()}.json`);
        showToast(`「${company}」の${companyProjects.length}件を保存しました`, 'success');
      });
    });

    // カードクリックでプロジェクト選択
    list.querySelectorAll('.project-card').forEach(card => {
      card.addEventListener('click', (e) => {
        if (e.target.closest('.delete-btn') || e.target.closest('.btn-save-proj')) return;
        selectProject(card.dataset.id);
      });
    });

    // プロジェクト保存ボタン
    list.querySelectorAll('.btn-save-proj').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const proj = appData.projects.find(p => p.id === btn.dataset.save);
        if (!proj) return;
        const exportData = { _type: 'bpi_project', _version: 1, project: proj };
        downloadJson(exportData, `bpi_project_${sanitizeFilename(proj.company)}_${sanitizeFilename(proj.name)}_${today()}.json`);
        showToast(`「${proj.name}」を保存しました`, 'success');
      });
    });

    // 削除ボタン
    list.querySelectorAll('.delete-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (confirm('このプロジェクトを削除しますか？')) {
          appData.projects = appData.projects.filter(p => p.id !== btn.dataset.delete);
          if (appData.currentProjectId === btn.dataset.delete) {
            appData.currentProjectId = null;
          }
          saveData(appData);
          renderDashboard();
          showToast('プロジェクトを削除しました');
        }
      });
    });
  }

  function selectProject(id) {
    appData.currentProjectId = id;
    saveData(appData);
    const proj = getCurrentProject();
    if (proj) {
      $('#currentProjectInfo').style.display = '';
      $('#projectBadge').textContent = `${proj.company} / ${proj.name}`;
      navigate('step' + (proj.currentStep || 1));
    }
  }

  // ==========================================
  // Step 1: 業務プロセスの洗い出し
  // ==========================================

  /** ムリ・ムダ・ムラ配列をタグHTMLに変換 */
  function renderMmmTags(mmm) {
    if (!mmm || !Array.isArray(mmm) || mmm.length === 0) return '<span style="color:var(--text-secondary)">-</span>';
    return mmm.map(m => {
      const cls = m === 'ムリ' ? 'mmm-tag-muri' : m === 'ムダ' ? 'mmm-tag-muda' : 'mmm-tag-mura';
      return `<span class="mmm-tag ${cls}">${escapeHtml(m)}</span>`;
    }).join(' ');
  }

  let step1Filter = 'all';

  function renderStep1() {
    const proj = getCurrentProject();
    if (!proj) return;
    if (!proj.tasks) proj.tasks = [];

    renderProcessTabs(proj);
    renderTaskTable(proj);
    renderStep1Stats(proj);
  }

  function renderProcessTabs(proj) {
    const tabs = $('#processTabs');
    const categories = appData.processCategories;
    const tasks = proj.tasks || [];

    let html = `<div class="process-tab ${step1Filter === 'all' ? 'active' : ''}" data-filter="all">
      すべて <span class="count">(${tasks.length})</span></div>`;

    categories.forEach(cat => {
      const count = tasks.filter(t => t.category === cat).length;
      html += `<div class="process-tab ${step1Filter === cat ? 'active' : ''}" data-filter="${escapeHtml(cat)}">
        ${escapeHtml(cat)} <span class="count">(${count})</span></div>`;
    });

    tabs.innerHTML = html;
    tabs.querySelectorAll('.process-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        step1Filter = tab.dataset.filter;
        renderStep1();
      });
    });
  }

  function renderTaskTable(proj) {
    const tbody = $('#taskTableBody');
    const noTasks = $('#noTasks');
    let tasks = proj.tasks || [];

    if (step1Filter !== 'all') {
      tasks = tasks.filter(t => t.category === step1Filter);
    }

    if (tasks.length === 0) {
      tbody.innerHTML = '';
      noTasks.style.display = '';
      return;
    }

    noTasks.style.display = 'none';
    tbody.innerHTML = tasks.map((t, i) => `
      <tr>
        <td>${i + 1}</td>
        <td><span style="font-size:11px;color:var(--text-secondary)">${escapeHtml(t.code)}</span></td>
        <td>${escapeHtml(t.category)}</td>
        <td><strong>${escapeHtml(t.content)}</strong></td>
        <td>${escapeHtml(t.person)}</td>
        <td>${escapeHtml(t.target)}</td>
        <td>${escapeHtml(t.method)}</td>
        <td>${t.timeRequired || '-'} ${escapeHtml(t.timeUnit || '')}</td>
        <td>${escapeHtml(t.freqType || '')} ${t.freqCount || ''}回</td>
        <td>${escapeHtml(t.tools)}</td>
        <td>${renderMmmTags(t.mmm)}</td>
        <td class="actions">
          <button class="btn-move" onclick="window.BPI.moveTask('${t.id}','up')" title="上に移動" ${i === 0 ? 'disabled' : ''}>▲</button>
          <button class="btn-move" onclick="window.BPI.moveTask('${t.id}','down')" title="下に移動" ${i === tasks.length - 1 ? 'disabled' : ''}>▼</button>
          <button class="btn btn-xs btn-outline" onclick="window.BPI.editTask('${t.id}')">編集</button>
          <button class="btn btn-xs btn-outline" onclick="window.BPI.deleteTask('${t.id}')" style="color:var(--danger)">削除</button>
        </td>
      </tr>
    `).join('');
  }

  function renderStep1Stats(proj) {
    const tasks = proj.tasks || [];
    const totalMonthly = tasks.reduce((sum, t) => sum + calcMonthlyTime(t), 0);
    const categories = [...new Set(tasks.map(t => t.category))];

    $('#step1Stats').innerHTML = `
      登録業務: <strong>${tasks.length}</strong>件
      ／ プロセス区分: <strong>${categories.length}</strong>種類
      ／ 月間合計時間: <strong>${Math.round(totalMonthly)}</strong>分
      （約<strong>${(totalMonthly / 60).toFixed(1)}</strong>時間）
    `;
  }

  function renderProcessMap() {
    const proj = getCurrentProject();
    if (!proj || !proj.tasks || proj.tasks.length === 0) {
      $('#processMapArea').innerHTML = '<div class="empty-state"><p>業務データを登録するとプロセスマップが表示されます。</p></div>';
      return;
    }

    const grouped = {};
    proj.tasks.forEach(t => {
      if (!grouped[t.category]) grouped[t.category] = [];
      grouped[t.category].push(t);
    });

    let html = '<div class="process-map">';
    Object.keys(grouped).forEach(cat => {
      html += `<div class="process-lane">
        <div class="lane-header">${escapeHtml(cat)}</div>
        <div class="lane-tasks">
          ${grouped[cat].map(t => `
            <div class="lane-task">
              <div>${escapeHtml(t.content)}</div>
              <div class="task-person">${escapeHtml(t.person)} / ${t.timeRequired || '?'}${escapeHtml(t.timeUnit || '分')}</div>
            </div>
          `).join('')}
        </div>
      </div>`;
    });
    html += '</div>';
    $('#processMapArea').innerHTML = html;
  }

  let editingTaskId = null;

  function openTaskModal(taskId) {
    editingTaskId = taskId || null;
    const proj = getCurrentProject();

    // カテゴリ選択肢
    const sel = $('#taskCategory');
    sel.innerHTML = appData.processCategories.map(c =>
      `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`
    ).join('');

    if (taskId) {
      const task = (proj.tasks || []).find(t => t.id === taskId);
      if (!task) return;
      $('#modalTaskTitle').textContent = '業務タスクの編集';
      sel.value = task.category;
      $('#taskCode').value = task.code;
      $('#taskContent').value = task.content || '';
      $('#taskPerson').value = task.person || '';
      $('#taskTarget').value = task.target || '';
      $('#taskMethod').value = task.method || '';
      $('#taskTime').value = task.timeRequired || '';
      $('#taskTimeUnit').value = task.timeUnit || '分';
      $('#taskFreqType').value = task.freqType || '日次';
      $('#taskFreqCount').value = task.freqCount || 1;
      $('#taskTools').value = task.tools || '';
      const mmm = task.mmm || [];
      $('#taskMuri').checked = mmm.includes('ムリ');
      $('#taskMuda').checked = mmm.includes('ムダ');
      $('#taskMura').checked = mmm.includes('ムラ');
    } else {
      $('#modalTaskTitle').textContent = '業務タスクの追加';
      $('#taskCode').value = generateTaskCode(sel.value);
      $('#taskContent').value = '';
      $('#taskPerson').value = '';
      $('#taskTarget').value = '';
      $('#taskMethod').value = '';
      $('#taskTime').value = '';
      $('#taskTimeUnit').value = '分';
      $('#taskFreqType').value = '日次';
      $('#taskFreqCount').value = 1;
      $('#taskTools').value = '';
      $('#taskMuri').checked = false;
      $('#taskMuda').checked = false;
      $('#taskMura').checked = false;
    }

    sel.addEventListener('change', () => {
      if (!editingTaskId) {
        $('#taskCode').value = generateTaskCode(sel.value);
      }
    });

    openModal('modalTask');
  }

  function saveTask() {
    const proj = getCurrentProject();
    if (!proj) return;
    if (!proj.tasks) proj.tasks = [];

    const content = $('#taskContent').value.trim();
    if (!content) { showToast('作業内容を入力してください', 'error'); return; }

    const taskData = {
      category: $('#taskCategory').value,
      code: $('#taskCode').value,
      content,
      person: $('#taskPerson').value.trim(),
      target: $('#taskTarget').value.trim(),
      method: $('#taskMethod').value.trim(),
      timeRequired: $('#taskTime').value,
      timeUnit: $('#taskTimeUnit').value,
      freqType: $('#taskFreqType').value,
      freqCount: $('#taskFreqCount').value,
      tools: $('#taskTools').value.trim(),
      notes: '',
      mmm: [
        ...$('#taskMuri').checked ? ['ムリ'] : [],
        ...$('#taskMuda').checked ? ['ムダ'] : [],
        ...$('#taskMura').checked ? ['ムラ'] : []
      ]
    };

    if (editingTaskId) {
      const idx = proj.tasks.findIndex(t => t.id === editingTaskId);
      if (idx >= 0) {
        proj.tasks[idx] = { ...proj.tasks[idx], ...taskData };
      }
      showToast('業務タスクを更新しました', 'success');
    } else {
      proj.tasks.push({ id: genId(), ...taskData, scores: null, problems: [], ecrs: null });
      showToast('業務タスクを追加しました', 'success');
    }

    saveData(appData);
    closeModal('modalTask');
    renderStep1();
  }

  function deleteTask(taskId) {
    if (!confirm('この業務タスクを削除しますか？')) return;
    const proj = getCurrentProject();
    if (!proj) return;
    proj.tasks = (proj.tasks || []).filter(t => t.id !== taskId);
    saveData(appData);
    renderStep1();
    showToast('削除しました');
  }

  // --- 業務の並び替え ---
  function moveTask(taskId, direction) {
    const proj = getCurrentProject();
    if (!proj || !proj.tasks) return;
    const allTasks = proj.tasks;

    if (step1Filter !== 'all') {
      // カテゴリフィルタ中：フィルタ内での移動を本体配列に反映
      const filtered = allTasks.filter(t => t.category === step1Filter);
      const fi = filtered.findIndex(t => t.id === taskId);
      if (fi < 0) return;
      const swapFi = direction === 'up' ? fi - 1 : fi + 1;
      if (swapFi < 0 || swapFi >= filtered.length) return;
      // 本体配列上のインデックスを取得して swap
      const ai = allTasks.indexOf(filtered[fi]);
      const bi = allTasks.indexOf(filtered[swapFi]);
      [allTasks[ai], allTasks[bi]] = [allTasks[bi], allTasks[ai]];
    } else {
      // 全件表示：配列内で単純 swap
      const idx = allTasks.findIndex(t => t.id === taskId);
      if (idx < 0) return;
      const swapIdx = direction === 'up' ? idx - 1 : idx + 1;
      if (swapIdx < 0 || swapIdx >= allTasks.length) return;
      [allTasks[idx], allTasks[swapIdx]] = [allTasks[swapIdx], allTasks[idx]];
    }
    saveData(appData);
    renderTaskTable(proj);
  }

  function renumberTaskCodes(proj) {
    if (!proj || !proj.tasks) return;
    const counters = {};
    proj.tasks.forEach(t => {
      const prefix = CODE_PREFIX[t.category] || 'TSK';
      counters[t.category] = (counters[t.category] || 0) + 1;
      t.code = `${prefix}-${String(counters[t.category]).padStart(3, '0')}`;
    });
    saveData(appData);
    renderTaskTable(proj);
    showToast('コードを再採番しました', 'success');
  }

  function importCSV() {
    const fileInput = $('#csvFileInput');
    fileInput.click();
    fileInput.onchange = () => {
      const file = fileInput.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          parseCSVAndImport(e.target.result);
        } catch (err) {
          showToast('CSVの読み込みに失敗しました: ' + err.message, 'error');
        }
      };
      reader.readAsText(file, 'UTF-8');
      fileInput.value = '';
    };
  }

  function parseCSVAndImport(csvText) {
    const proj = getCurrentProject();
    if (!proj) return;
    if (!proj.tasks) proj.tasks = [];

    const lines = csvText.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) { showToast('データが不足しています', 'error'); return; }

    const headers = lines[0].split(/[,\t]/).map(h => h.replace(/"/g, '').trim());

    // 列マッピング（柔軟に対応）
    const mapping = {};
    const maps = {
      category: ['プロセス区分', '区分', 'カテゴリ', 'category'],
      content: ['作業内容', '内容', '何を', 'content', '業務内容'],
      person: ['担当者', '誰が', 'person', '担当'],
      target: ['対象', '誰に', 'target'],
      method: ['方法', 'どうやって', 'method'],
      timeRequired: ['所要時間', '時間', 'time'],
      freqType: ['頻度', 'frequency'],
      tools: ['ツール', '使用ツール', 'tools', '書類'],
      notes: ['備考', '問題点', '気づき', 'notes', 'メモ', 'ムリ・ムダ・ムラ', '3M']
    };

    headers.forEach((h, i) => {
      for (const [key, keywords] of Object.entries(maps)) {
        if (keywords.some(k => h.includes(k))) {
          mapping[key] = i;
          break;
        }
      }
    });

    let imported = 0;
    for (let i = 1; i < lines.length; i++) {
      const cols = lines[i].split(/[,\t]/).map(c => c.replace(/"/g, '').trim());
      if (cols.length < 2) continue;

      const task = {
        id: genId(),
        category: (mapping.category !== undefined ? cols[mapping.category] : '') || 'その他',
        code: '',
        content: (mapping.content !== undefined ? cols[mapping.content] : cols[0]) || '',
        person: mapping.person !== undefined ? cols[mapping.person] : '',
        target: mapping.target !== undefined ? cols[mapping.target] : '',
        method: mapping.method !== undefined ? cols[mapping.method] : '',
        timeRequired: mapping.timeRequired !== undefined ? cols[mapping.timeRequired] : '',
        timeUnit: '分',
        freqType: mapping.freqType !== undefined ? cols[mapping.freqType] : '月次',
        freqCount: '1',
        tools: mapping.tools !== undefined ? cols[mapping.tools] : '',
        notes: mapping.notes !== undefined ? cols[mapping.notes] : '',
        mmm: [],
        scores: null,
        problems: [],
        ecrs: null
      };

      // CSV備考欄からムリ・ムダ・ムラを解析
      if (task.notes) {
        const mmmValues = ['ムリ', 'ムダ', 'ムラ'];
        const parsed = mmmValues.filter(v => task.notes.includes(v));
        if (parsed.length > 0) { task.mmm = parsed; task.notes = ''; }
      }

      if (!task.content) continue;
      task.code = generateTaskCode(task.category);

      // 新しいカテゴリがあれば追加
      if (task.category && !appData.processCategories.includes(task.category)) {
        appData.processCategories.push(task.category);
      }

      proj.tasks.push(task);
      imported++;
    }

    saveData(appData);
    renderStep1();
    showToast(`${imported}件の業務をインポートしました`, 'success');
  }

  // ==========================================
  // Step 2: 問題点の可視化と分析
  // ==========================================
  function renderStep2() {
    const proj = getCurrentProject();
    if (!proj || !proj.tasks) return;
    renderScoringList(proj);
  }

  function renderScoringList(proj) {
    const container = $('#scoringList');
    const tasks = proj.tasks || [];

    if (tasks.length === 0) {
      container.innerHTML = '<div class="empty-state"><p>Step 1で業務を登録してください。</p></div>';
      return;
    }

    container.innerHTML = tasks.map(t => {
      const scores = t.scores || { timeImpact: 1, qualityImpact: 1, frequency: 1, difficulty: 5 };
      const total = calcTotalScore(scores).toFixed(1);
      const problems = t.problems || [];

      return `
        <div class="scoring-card" data-task-id="${t.id}">
          <div class="scoring-card-header">
            <h4>${escapeHtml(t.content)}</h4>
            <span class="task-code">${escapeHtml(t.code)} / ${escapeHtml(t.category)}</span>
          </div>
          <div style="font-size:12px;color:var(--text-secondary);margin-bottom:8px;">
            ${escapeHtml(t.person)} / ${t.timeRequired || '-'}${escapeHtml(t.timeUnit || '分')} / ${escapeHtml(t.freqType || '')}${t.freqCount || ''}回
            ${t.notes ? `<br>備考: ${escapeHtml(t.notes)}` : ''}
          </div>
          <div style="margin-bottom:8px;">
            <label style="font-size:12px;font-weight:600;color:var(--text-secondary);">問題カテゴリ（複数選択可）</label>
          </div>
          <div class="problem-tags">
            ${PROBLEM_CATEGORIES.map(pc => `
              <div class="problem-tag ${problems.includes(pc.id) ? 'active' : ''}"
                   data-task="${t.id}" data-problem="${pc.id}"
                   title="${escapeHtml(pc.desc)}">
                ${escapeHtml(pc.name)}
              </div>
            `).join('')}
          </div>
          <div class="score-inputs">
            <div class="score-input">
              <label>時間影響度</label>
              <div class="score-slider">
                <input type="range" min="1" max="5" value="${scores.timeImpact}"
                       data-task="${t.id}" data-score="timeImpact">
                <span class="score-value">${scores.timeImpact}</span>
              </div>
            </div>
            <div class="score-input">
              <label>品質影響度</label>
              <div class="score-slider">
                <input type="range" min="1" max="5" value="${scores.qualityImpact}"
                       data-task="${t.id}" data-score="qualityImpact">
                <span class="score-value">${scores.qualityImpact}</span>
              </div>
            </div>
            <div class="score-input">
              <label>頻度</label>
              <div class="score-slider">
                <input type="range" min="1" max="5" value="${scores.frequency}"
                       data-task="${t.id}" data-score="frequency">
                <span class="score-value">${scores.frequency}</span>
              </div>
            </div>
            <div class="score-input">
              <label>改善難易度（高=容易）</label>
              <div class="score-slider">
                <input type="range" min="1" max="5" value="${scores.difficulty}"
                       data-task="${t.id}" data-score="difficulty">
                <span class="score-value">${scores.difficulty}</span>
              </div>
            </div>
          </div>
          <div class="total-score">総合優先度スコア: ${total}</div>
        </div>
      `;
    }).join('');

    // スコアスライダーのイベント
    container.querySelectorAll('input[type="range"]').forEach(slider => {
      slider.addEventListener('input', (e) => {
        const taskId = e.target.dataset.task;
        const scoreKey = e.target.dataset.score;
        const val = parseInt(e.target.value);

        e.target.nextElementSibling.textContent = val;

        const task = proj.tasks.find(t => t.id === taskId);
        if (task) {
          if (!task.scores) task.scores = { timeImpact: 1, qualityImpact: 1, frequency: 1, difficulty: 5 };
          task.scores[scoreKey] = val;
          const total = calcTotalScore(task.scores).toFixed(1);
          const card = e.target.closest('.scoring-card');
          card.querySelector('.total-score').textContent = `総合優先度スコア: ${total}`;
          saveData(appData);
        }
      });
    });

    // 問題タグのイベント
    container.querySelectorAll('.problem-tag').forEach(tag => {
      tag.addEventListener('click', () => {
        const taskId = tag.dataset.task;
        const problemId = tag.dataset.problem;
        const task = proj.tasks.find(t => t.id === taskId);
        if (!task) return;
        if (!task.problems) task.problems = [];

        const idx = task.problems.indexOf(problemId);
        if (idx >= 0) {
          task.problems.splice(idx, 1);
          tag.classList.remove('active');
        } else {
          task.problems.push(problemId);
          tag.classList.add('active');
        }
        saveData(appData);
      });
    });
  }

  function renderAnalysisDashboard() {
    const proj = getCurrentProject();
    if (!proj || !proj.tasks || proj.tasks.length === 0) return;

    renderParetoChart(proj);
    renderHeatmap(proj);
    renderTimeAnalysis(proj);
    renderBubbleChart(proj);
  }

  let chartInstances = {};

  function destroyChart(key) {
    if (chartInstances[key]) {
      chartInstances[key].destroy();
      delete chartInstances[key];
    }
  }

  function renderParetoChart(proj) {
    destroyChart('pareto');
    const tasks = proj.tasks || [];
    const categoryCounts = {};

    tasks.forEach(t => {
      (t.problems || []).forEach(p => {
        const cat = PROBLEM_CATEGORIES.find(pc => pc.id === p);
        if (cat) categoryCounts[cat.name] = (categoryCounts[cat.name] || 0) + 1;
      });
    });

    const sorted = Object.entries(categoryCounts).sort((a, b) => b[1] - a[1]);
    if (sorted.length === 0) return;

    const labels = sorted.map(s => s[0]);
    const values = sorted.map(s => s[1]);
    const total = values.reduce((a, b) => a + b, 0);
    let cum = 0;
    const cumPercent = values.map(v => { cum += v; return Math.round((cum / total) * 100); });

    const ctx = $('#chartPareto').getContext('2d');
    chartInstances.pareto = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: '件数', data: values, backgroundColor: 'rgba(37,99,235,0.6)', yAxisID: 'y' },
          { label: '累積%', data: cumPercent, type: 'line', borderColor: '#DC2626', yAxisID: 'y1', pointRadius: 4 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        scales: {
          y: { beginAtZero: true, position: 'left', title: { display: true, text: '件数' } },
          y1: { beginAtZero: true, max: 100, position: 'right', title: { display: true, text: '累積%' }, grid: { drawOnChartArea: false } }
        }
      }
    });
  }

  function renderHeatmap(proj) {
    const container = $('#heatmapContainer');
    const tasks = proj.tasks || [];
    const categories = [...new Set(tasks.map(t => t.category))];

    if (categories.length === 0) { container.innerHTML = ''; return; }

    const matrix = {};
    categories.forEach(cat => {
      matrix[cat] = {};
      PROBLEM_CATEGORIES.forEach(pc => { matrix[cat][pc.id] = 0; });
    });

    tasks.forEach(t => {
      (t.problems || []).forEach(p => {
        if (matrix[t.category]) matrix[t.category][p]++;
      });
    });

    const maxVal = Math.max(1, ...Object.values(matrix).flatMap(r => Object.values(r)));

    let html = '<table class="heatmap-table"><thead><tr><th></th>';
    PROBLEM_CATEGORIES.forEach(pc => { html += `<th>${escapeHtml(pc.name)}</th>`; });
    html += '</tr></thead><tbody>';

    categories.forEach(cat => {
      html += `<tr><th style="text-align:left">${escapeHtml(cat)}</th>`;
      PROBLEM_CATEGORIES.forEach(pc => {
        const v = matrix[cat][pc.id];
        const level = Math.min(5, Math.ceil((v / maxVal) * 5));
        html += `<td class="heat-${v > 0 ? level : 0}">${v || ''}</td>`;
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
  }

  function renderTimeAnalysis(proj) {
    destroyChart('timeAnalysis');
    const tasks = proj.tasks || [];
    const categories = [...new Set(tasks.map(t => t.category))];
    if (categories.length === 0) return;

    const valueAdded = [];
    const nonValueAdded = [];

    categories.forEach(cat => {
      const catTasks = tasks.filter(t => t.category === cat);
      let va = 0, nva = 0;
      catTasks.forEach(t => {
        const monthTime = calcMonthlyTime(t);
        const hasProblems = (t.problems || []).length > 0;
        if (hasProblems) nva += monthTime;
        else va += monthTime;
      });
      valueAdded.push(Math.round(va));
      nonValueAdded.push(Math.round(nva));
    });

    const ctx = $('#chartTimeAnalysis').getContext('2d');
    chartInstances.timeAnalysis = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: categories,
        datasets: [
          { label: '付加価値作業', data: valueAdded, backgroundColor: 'rgba(16,185,129,0.6)' },
          { label: '非付加価値作業', data: nonValueAdded, backgroundColor: 'rgba(239,68,68,0.6)' }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { title: { display: false } },
        scales: {
          x: { stacked: true },
          y: { stacked: true, title: { display: true, text: '月間時間（分）' } }
        }
      }
    });
  }

  function renderBubbleChart(proj) {
    destroyChart('bubble');
    const tasks = proj.tasks || [];
    const data = tasks.filter(t => t.scores).map(t => {
      const s = t.scores;
      const total = calcTotalScore(s);
      return {
        x: total,
        y: s.difficulty,
        r: Math.max(5, (calcMonthlyTime(t) / 10)),
        label: t.content
      };
    });

    if (data.length === 0) return;

    const ctx = $('#chartBubble').getContext('2d');
    chartInstances.bubble = new Chart(ctx, {
      type: 'bubble',
      data: {
        datasets: [{
          label: '業務',
          data,
          backgroundColor: 'rgba(37,99,235,0.4)',
          borderColor: 'rgba(37,99,235,0.8)',
          borderWidth: 1
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        scales: {
          x: { title: { display: true, text: '改善効果（優先度スコア）' }, min: 0, max: 5 },
          y: { title: { display: true, text: '実施難易度（高=容易）' }, min: 0, max: 6 }
        },
        plugins: {
          tooltip: {
            callbacks: {
              label: (ctx) => ctx.raw.label || ''
            }
          }
        }
      }
    });
  }

  // ==========================================
  // Step 3: ECRS改善対象の選定
  // ==========================================
  function renderStep3() {
    const proj = getCurrentProject();
    if (!proj || !proj.tasks) return;
    renderECRSWizard(proj);
  }

  function renderECRSWizard(proj) {
    const area = $('#ecrsWizardArea');
    const tasks = proj.tasks || [];

    // 優先度スコアで降順ソート
    const sorted = [...tasks].sort((a, b) => {
      const sa = a.scores ? calcTotalScore(a.scores) : 0;
      const sb = b.scores ? calcTotalScore(b.scores) : 0;
      return sb - sa;
    });

    if (sorted.length === 0) {
      area.innerHTML = '<div class="empty-state"><p>Step 1で業務を登録してください。</p></div>';
      return;
    }

    area.innerHTML = sorted.map(t => {
      const ecrs = t.ecrs || {};
      const totalScore = t.scores ? calcTotalScore(t.scores).toFixed(1) : '-';

      return `
        <div class="ecrs-wizard-task" data-task-id="${t.id}">
          <div class="ecrs-task-info">
            <h4>${escapeHtml(t.content)}</h4>
            <p>${escapeHtml(t.category)} / ${escapeHtml(t.person)} / 優先度: ${totalScore}</p>
          </div>
          <div class="ecrs-questions">
            ${ECRS_DEFINITIONS.map(def => `
              <div class="ecrs-question ecrs-${def.color}">
                <h5>
                  <span class="ecrs-label ecrs-label-${def.color}">${def.key}</span>
                  ${escapeHtml(def.name)}
                </h5>
                <p>${escapeHtml(def.questions[0])}</p>
                <div class="ecrs-options">
                  <button class="ecrs-option ${ecrs[def.key] === 'yes' ? 'selected-yes' : ''}"
                          data-task="${t.id}" data-ecrs="${def.key}" data-answer="yes">
                    はい（${def.key}候補）
                  </button>
                  <button class="ecrs-option ${ecrs[def.key] === 'no' ? 'selected-no' : ''}"
                          data-task="${t.id}" data-ecrs="${def.key}" data-answer="no">
                    いいえ
                  </button>
                  <button class="ecrs-option ${ecrs[def.key] === 'maybe' ? 'selected' : ''}"
                          data-task="${t.id}" data-ecrs="${def.key}" data-answer="maybe">
                    要検討
                  </button>
                </div>
              </div>
            `).join('')}
          </div>
        </div>
      `;
    }).join('');

    // ECRS回答イベント
    area.querySelectorAll('.ecrs-option').forEach(btn => {
      btn.addEventListener('click', () => {
        const taskId = btn.dataset.task;
        const ecrsKey = btn.dataset.ecrs;
        const answer = btn.dataset.answer;

        const task = proj.tasks.find(t => t.id === taskId);
        if (!task) return;
        if (!task.ecrs) task.ecrs = {};
        task.ecrs[ecrsKey] = answer;

        // UIリセット
        const parent = btn.closest('.ecrs-options');
        parent.querySelectorAll('.ecrs-option').forEach(o => {
          o.className = 'ecrs-option';
        });
        if (answer === 'yes') btn.classList.add('selected-yes');
        else if (answer === 'no') btn.classList.add('selected-no');
        else btn.classList.add('selected');

        saveData(appData);
      });
    });
  }

  function renderCandidatesList() {
    const proj = getCurrentProject();
    if (!proj || !proj.tasks) return;

    const container = $('#candidatesList');
    const candidates = [];

    proj.tasks.forEach(t => {
      if (!t.ecrs) return;
      const ecrsEntries = [];
      ECRS_DEFINITIONS.forEach(def => {
        if (t.ecrs[def.key] === 'yes' || t.ecrs[def.key] === 'maybe') {
          ecrsEntries.push({
            task: t,
            ecrsKey: def.key,
            ecrsName: def.name,
            ecrsColor: def.color,
            answer: t.ecrs[def.key],
            score: t.scores ? calcTotalScore(t.scores) : 0
          });
        }
      });
      candidates.push(...ecrsEntries);
    });

    // スコアでソート
    candidates.sort((a, b) => b.score - a.score);

    if (candidates.length === 0) {
      container.innerHTML = '<div class="empty-state"><p>ECRSウィザードで判定を行ってください。</p></div>';
      return;
    }

    // ECRS区分でグルーピング
    const grouped = { E: [], C: [], R: [], S: [] };
    candidates.forEach(c => {
      if (grouped[c.ecrsKey]) grouped[c.ecrsKey].push(c);
    });

    let html = '';
    ECRS_DEFINITIONS.forEach(def => {
      const items = grouped[def.key];
      if (items.length === 0) return;
      html += `<h3 style="margin:16px 0 8px;display:flex;align-items:center;gap:8px;">
        <span class="ecrs-label ecrs-label-${def.color}">${def.key}</span> ${escapeHtml(def.name)}（${items.length}件）
      </h3>`;
      html += '<div class="candidates-grid">';
      items.forEach(c => {
        html += `
          <div class="candidate-card ecrs-${c.ecrsColor}">
            <h4>${escapeHtml(c.task.content)}</h4>
            <div class="candidate-meta">${escapeHtml(c.task.category)} / ${escapeHtml(c.task.person)} / スコア: ${c.score.toFixed(1)}</div>
            <div style="margin-top:8px;">
              <button class="btn btn-sm btn-primary" onclick="window.BPI.openImprovementModal('${c.task.id}', '${c.ecrsKey}')">
                改善計画を作成
              </button>
            </div>
          </div>
        `;
      });
      html += '</div>';
    });

    container.innerHTML = html;
  }

  // ==========================================
  // Step 4: 改善策の検討
  // ==========================================
  function renderStep4() {
    const proj = getCurrentProject();
    if (!proj) return;
    renderImprovementPlans(proj);
  }

  function renderImprovementPlans(proj) {
    const container = $('#improvementPlans');
    if (!proj.improvements) proj.improvements = [];

    // 改善候補があるタスクを表示
    const candidates = [];
    (proj.tasks || []).forEach(t => {
      if (!t.ecrs) return;
      ECRS_DEFINITIONS.forEach(def => {
        if (t.ecrs[def.key] === 'yes' || t.ecrs[def.key] === 'maybe') {
          const existing = proj.improvements.find(imp => imp.taskId === t.id && imp.ecrsKey === def.key);
          candidates.push({ task: t, ecrsKey: def.key, ecrsName: def.name, ecrsColor: def.color, improvement: existing });
        }
      });
    });

    if (candidates.length === 0) {
      container.innerHTML = '<div class="empty-state"><p>Step 3でECRS判定を行い、改善候補を特定してください。</p></div>';
      return;
    }

    let html = '<div class="plan-cards">';
    candidates.forEach(c => {
      const imp = c.improvement;
      const status = imp ? (imp.status || 'pending') : 'pending';
      const statusText = { pending: '未着手', progress: '進行中', done: '完了' }[status];
      const statusClass = { pending: 'status-pending', progress: 'status-progress', done: 'status-done' }[status];

      // 問題カテゴリに応じたテンプレート提案
      const problemIds = c.task.problems || [];
      const templates = [];
      problemIds.forEach(pid => {
        if (IMPROVEMENT_TEMPLATES[pid]) {
          templates.push(...IMPROVEMENT_TEMPLATES[pid]);
        }
      });
      const uniqueTemplates = [...new Set(templates)].slice(0, 3);

      html += `
        <div class="plan-card">
          <div class="plan-card-header">
            <h4>
              <span class="ecrs-label ecrs-label-${c.ecrsColor}" style="margin-right:6px;">${c.ecrsKey}</span>
              ${escapeHtml(c.task.content)}
            </h4>
            <span class="plan-status ${statusClass}">${statusText}</span>
          </div>
          ${imp ? `
            <div class="plan-details">
              <div class="plan-detail-item"><label>改善内容</label><strong>${escapeHtml(imp.content)}</strong></div>
              <div class="plan-detail-item"><label>手段</label><strong>${escapeHtml(imp.method || '-')}</strong></div>
              <div class="plan-detail-item"><label>担当</label><strong>${escapeHtml(imp.person || '-')}</strong></div>
              <div class="plan-detail-item"><label>期間</label><strong>${imp.startDate || '?'} 〜 ${imp.endDate || '?'}</strong></div>
              <div class="plan-detail-item"><label>期待効果</label><strong>${escapeHtml(imp.effectQuantity || '-')}</strong></div>
            </div>
            <div style="margin-top:8px;">
              <button class="btn btn-xs btn-outline" onclick="window.BPI.openImprovementModal('${c.task.id}', '${c.ecrsKey}')">編集</button>
              <select onchange="window.BPI.updateImpStatus('${c.task.id}','${c.ecrsKey}',this.value)" style="font-size:12px;padding:3px 6px;">
                <option value="pending" ${status==='pending'?'selected':''}>未着手</option>
                <option value="progress" ${status==='progress'?'selected':''}>進行中</option>
                <option value="done" ${status==='done'?'selected':''}>完了</option>
              </select>
            </div>
          ` : `
            <div style="margin:8px 0;">
              ${uniqueTemplates.length > 0 ? `
                <div style="font-size:12px;color:var(--text-secondary);margin-bottom:6px;">改善のヒント:</div>
                <div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:8px;">
                  ${uniqueTemplates.map(t => `<span style="font-size:11px;background:var(--warning-light);padding:2px 8px;border-radius:10px;">${escapeHtml(t)}</span>`).join('')}
                </div>
              ` : ''}
              <button class="btn btn-sm btn-primary" onclick="window.BPI.openImprovementModal('${c.task.id}', '${c.ecrsKey}')">
                改善計画を作成
              </button>
            </div>
          `}
        </div>
      `;
    });
    html += '</div>';
    container.innerHTML = html;
  }

  let editingImpTaskId = null;
  let editingImpEcrsKey = null;

  function openImprovementModal(taskId, ecrsKey) {
    editingImpTaskId = taskId;
    editingImpEcrsKey = ecrsKey;
    const proj = getCurrentProject();
    if (!proj) return;

    const task = proj.tasks.find(t => t.id === taskId);
    if (!task) return;

    const ecrsDef = ECRS_DEFINITIONS.find(d => d.key === ecrsKey);
    const existing = (proj.improvements || []).find(imp => imp.taskId === taskId && imp.ecrsKey === ecrsKey);

    $('#impTaskName').value = task.content;
    $('#impECRS').value = ecrsDef ? ecrsDef.name : ecrsKey;
    $('#impScore').value = task.scores ? calcTotalScore(task.scores).toFixed(1) : '-';

    if (existing) {
      $('#modalImprovementTitle').textContent = '改善アクションプランの編集';
      $('#impContent').value = existing.content || '';
      $('#impMethod').value = existing.method || '';
      $('#impPerson').value = existing.person || '';
      $('#impResources').value = existing.resources || '';
      $('#impStartDate').value = existing.startDate || '';
      $('#impEndDate').value = existing.endDate || '';
      $('#impEffectQuantity').value = existing.effectQuantity || '';
      $('#impEffectQuality').value = existing.effectQuality || '';
    } else {
      $('#modalImprovementTitle').textContent = '改善アクションプラン作成';
      $('#impContent').value = '';
      $('#impMethod').value = '';
      $('#impPerson').value = '';
      $('#impResources').value = '';
      $('#impStartDate').value = '';
      $('#impEndDate').value = '';
      $('#impEffectQuantity').value = '';
      $('#impEffectQuality').value = '';
    }

    openModal('modalImprovement');
  }

  function saveImprovement() {
    const proj = getCurrentProject();
    if (!proj) return;
    if (!proj.improvements) proj.improvements = [];

    const content = $('#impContent').value.trim();
    if (!content) { showToast('改善内容を入力してください', 'error'); return; }

    const impData = {
      taskId: editingImpTaskId,
      ecrsKey: editingImpEcrsKey,
      content,
      method: $('#impMethod').value,
      person: $('#impPerson').value.trim(),
      resources: $('#impResources').value.trim(),
      startDate: $('#impStartDate').value,
      endDate: $('#impEndDate').value,
      effectQuantity: $('#impEffectQuantity').value.trim(),
      effectQuality: $('#impEffectQuality').value.trim(),
      status: 'pending'
    };

    const idx = proj.improvements.findIndex(imp => imp.taskId === editingImpTaskId && imp.ecrsKey === editingImpEcrsKey);
    if (idx >= 0) {
      impData.status = proj.improvements[idx].status;
      proj.improvements[idx] = impData;
    } else {
      proj.improvements.push(impData);
    }

    saveData(appData);
    closeModal('modalImprovement');
    renderStep4();
    showToast('改善計画を保存しました', 'success');
  }

  function updateImpStatus(taskId, ecrsKey, status) {
    const proj = getCurrentProject();
    if (!proj) return;
    const imp = (proj.improvements || []).find(i => i.taskId === taskId && i.ecrsKey === ecrsKey);
    if (imp) {
      imp.status = status;
      saveData(appData);
      renderStep4();
    }
  }

  function renderGanttChart() {
    const proj = getCurrentProject();
    if (!proj || !proj.improvements) return;

    const container = $('#ganttChart');
    const imps = proj.improvements.filter(i => i.startDate && i.endDate);

    if (imps.length === 0) {
      container.innerHTML = '<div class="empty-state"><p>開始日・完了日が設定された改善計画がありません。</p></div>';
      return;
    }

    // 月のリストを生成
    const allDates = imps.flatMap(i => [new Date(i.startDate), new Date(i.endDate)]);
    const minDate = new Date(Math.min(...allDates));
    const maxDate = new Date(Math.max(...allDates));

    const months = [];
    const cur = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
    while (cur <= maxDate) {
      months.push(`${cur.getFullYear()}/${cur.getMonth() + 1}`);
      cur.setMonth(cur.getMonth() + 1);
    }
    if (months.length === 0) months.push(`${minDate.getFullYear()}/${minDate.getMonth() + 1}`);

    let html = '<table class="gantt-table"><thead><tr><th style="min-width:200px;text-align:left;">改善項目</th><th>状態</th>';
    months.forEach(m => { html += `<th>${m}</th>`; });
    html += '</tr></thead><tbody>';

    imps.forEach(imp => {
      const task = (proj.tasks || []).find(t => t.id === imp.taskId);
      const taskName = task ? task.content : '不明';
      const ecrsDef = ECRS_DEFINITIONS.find(d => d.key === imp.ecrsKey);
      const barClass = ecrsDef ? `gantt-bar-${ecrsDef.color}` : '';
      const statusText = { pending: '未着手', progress: '進行中', done: '完了' }[imp.status] || '未着手';

      const start = new Date(imp.startDate);
      const end = new Date(imp.endDate);

      html += `<tr><td style="font-size:12px;">${escapeHtml(taskName)}</td><td style="text-align:center;font-size:11px;">${statusText}</td>`;

      months.forEach(m => {
        const [y, mo] = m.split('/').map(Number);
        const monthStart = new Date(y, mo - 1, 1);
        const monthEnd = new Date(y, mo, 0);

        const isInRange = start <= monthEnd && end >= monthStart;
        html += `<td>${isInRange ? `<div class="gantt-bar ${barClass}"></div>` : ''}</td>`;
      });

      html += '</tr>';
    });

    html += '</tbody></table>';
    container.innerHTML = html;
  }

  // ==========================================
  // Step 5: 効果検証と振り返り
  // ==========================================
  function renderStep5() {
    const proj = getCurrentProject();
    if (!proj) return;
    renderMeasurements(proj);
  }

  function renderMeasurements(proj) {
    const container = $('#measurementArea');
    const imps = (proj.improvements || []).filter(i => i.status === 'done' || i.status === 'progress');

    if (imps.length === 0) {
      container.innerHTML = '<div class="empty-state"><p>改善計画が進行中または完了になると、効果測定が可能になります。</p></div>';
      return;
    }

    if (!proj.measurements) proj.measurements = [];

    let html = '<div class="measure-cards">';
    imps.forEach(imp => {
      const task = (proj.tasks || []).find(t => t.id === imp.taskId);
      if (!task) return;

      const measurement = proj.measurements.find(m => m.taskId === imp.taskId && m.ecrsKey === imp.ecrsKey);
      const beforeTime = parseFloat(task.timeRequired) || 0;
      const afterTime = measurement ? parseFloat(measurement.afterTime) : null;
      const reduction = afterTime !== null ? Math.round(((beforeTime - afterTime) / beforeTime) * 100) : null;

      html += `
        <div class="measure-card">
          <div style="display:flex;justify-content:space-between;align-items:center;">
            <h4>${escapeHtml(task.content)}</h4>
            <button class="btn btn-sm btn-outline" onclick="window.BPI.openMeasureModal('${imp.taskId}', '${imp.ecrsKey}')">
              ${measurement ? '測定データ編集' : '効果測定入力'}
            </button>
          </div>
          <div style="font-size:12px;color:var(--text-secondary);margin:4px 0;">改善内容: ${escapeHtml(imp.content)}</div>
          ${afterTime !== null ? `
            <div class="measure-comparison">
              <div class="measure-value">
                <div class="label">改善前</div>
                <div class="number before">${beforeTime}</div>
                <div class="label">${escapeHtml(task.timeUnit || '分')}</div>
              </div>
              <div class="measure-arrow">&rarr;</div>
              <div class="measure-value">
                <div class="label">改善後</div>
                <div class="number after">${afterTime}</div>
                <div class="label">${escapeHtml(task.timeUnit || '分')}</div>
              </div>
              <div class="measure-reduction">${reduction >= 0 ? '-' : '+'}${Math.abs(reduction)}%</div>
            </div>
            ${measurement.comment ? `<div style="font-size:12px;color:var(--text-secondary);">コメント: ${escapeHtml(measurement.comment)}</div>` : ''}
          ` : `<div style="margin:12px 0;font-size:13px;color:var(--text-light);">まだ効果測定データが入力されていません。</div>`}
        </div>
      `;
    });
    html += '</div>';
    container.innerHTML = html;
  }

  let measureTaskId = null;
  let measureEcrsKey = null;

  function openMeasureModal(taskId, ecrsKey) {
    measureTaskId = taskId;
    measureEcrsKey = ecrsKey;
    const proj = getCurrentProject();
    if (!proj) return;

    const task = proj.tasks.find(t => t.id === taskId);
    if (!task) return;

    const existing = (proj.measurements || []).find(m => m.taskId === taskId && m.ecrsKey === ecrsKey);

    $('#measureTaskName').value = task.content;
    $('#measureUnit').textContent = task.timeUnit || '分';

    if (existing) {
      $('#measureAfterTime').value = existing.afterTime || '';
      $('#measureAchievement').value = existing.achievement || '';
      $('#measureDate').value = existing.date || '';
      $('#measureComment').value = existing.comment || '';
    } else {
      $('#measureAfterTime').value = '';
      $('#measureAchievement').value = '';
      $('#measureDate').value = new Date().toISOString().split('T')[0];
      $('#measureComment').value = '';
    }

    openModal('modalMeasure');
  }

  function saveMeasure() {
    const proj = getCurrentProject();
    if (!proj) return;
    if (!proj.measurements) proj.measurements = [];

    const data = {
      taskId: measureTaskId,
      ecrsKey: measureEcrsKey,
      afterTime: $('#measureAfterTime').value,
      achievement: $('#measureAchievement').value,
      date: $('#measureDate').value,
      comment: $('#measureComment').value.trim()
    };

    const idx = proj.measurements.findIndex(m => m.taskId === measureTaskId && m.ecrsKey === measureEcrsKey);
    if (idx >= 0) proj.measurements[idx] = data;
    else proj.measurements.push(data);

    saveData(appData);
    closeModal('modalMeasure');
    renderStep5();
    showToast('効果測定データを保存しました', 'success');
  }

  // ==========================================
  // レポートパネル（Step5サブビュー内のサマリ＋ボタン）
  // ==========================================
  function renderReportPanel() {
    const proj = getCurrentProject();
    if (!proj) return;

    const container = $('#reportArea');
    const tasks = proj.tasks || [];
    const imps = proj.improvements || [];
    const measures = proj.measurements || [];

    const totalTasks = tasks.length;
    const improvedCount = imps.length;
    const doneCount = imps.filter(i => i.status === 'done').length;

    let totalTimeSaved = 0;
    measures.forEach(m => {
      const task = tasks.find(t => t.id === m.taskId);
      if (task) {
        const before = parseFloat(task.timeRequired) || 0;
        const after = parseFloat(m.afterTime) || 0;
        totalTimeSaved += (before - after);
      }
    });

    const ecrsBreakdown = { E: 0, C: 0, R: 0, S: 0 };
    imps.forEach(i => { if (ecrsBreakdown[i.ecrsKey] !== undefined) ecrsBreakdown[i.ecrsKey]++; });

    let html = `
      <div class="report-summary">
        <h3 style="margin-bottom:16px;">改善活動サマリ</h3>
        <div class="report-kpis">
          <div class="kpi-card">
            <div class="kpi-value">${totalTasks}</div>
            <div class="kpi-label">総業務プロセス数</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-value">${improvedCount}</div>
            <div class="kpi-label">改善対象件数</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-value">${doneCount}</div>
            <div class="kpi-label">改善完了件数</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-value">${Math.round(totalTimeSaved)}分</div>
            <div class="kpi-label">削減時間（1回あたり）</div>
          </div>
        </div>
        <h4 style="margin-bottom:12px;">ECRS区分別 改善件数</h4>
        <div style="display:flex;gap:12px;margin-bottom:20px;">
          ${ECRS_DEFINITIONS.map(def => `
            <div style="flex:1;text-align:center;padding:12px;background:var(--bg);border-radius:var(--radius);">
              <span class="ecrs-label ecrs-label-${def.color}" style="font-size:14px;">${def.key}</span>
              <div style="font-size:20px;font-weight:700;margin-top:4px;">${ecrsBreakdown[def.key]}</div>
              <div style="font-size:11px;color:var(--text-secondary);">${def.name.split('（')[0]}</div>
            </div>
          `).join('')}
        </div>
        ${measures.length > 0 ? `
          <h4 style="margin-bottom:12px;">効果測定結果</h4>
          <table class="data-table" style="margin-bottom:0;">
            <thead><tr><th>業務</th><th>改善前</th><th>改善後</th><th>削減率</th></tr></thead>
            <tbody>
              ${measures.map(m => {
                const task = tasks.find(t => t.id === m.taskId);
                if (!task) return '';
                const before = parseFloat(task.timeRequired) || 0;
                const after = parseFloat(m.afterTime) || 0;
                const pct = before > 0 ? Math.round(((before - after) / before) * 100) : 0;
                return `<tr><td>${escapeHtml(task.content)}</td><td>${before}${escapeHtml(task.timeUnit||'分')}</td><td>${after}${escapeHtml(task.timeUnit||'分')}</td><td style="color:${pct>=0?'var(--success)':'var(--danger)'}">${pct>=0?'-':''}${Math.abs(pct)}%</td></tr>`;
              }).join('')}
            </tbody>
          </table>
        ` : ''}
        <div style="margin-top:24px;display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">
          <button class="btn btn-primary" id="btnPrintReport">&#128424; 印刷用レポートを生成</button>
          <button class="btn btn-outline" id="btnExcelExport">&#128202; Excelエクスポート</button>
        </div>
      </div>
    `;
    container.innerHTML = html;

    $('#btnPrintReport').addEventListener('click', () => {
      buildPrintReport();
      setTimeout(() => window.print(), 400);
    });
    $('#btnExcelExport').addEventListener('click', exportExcel);
  }

  // ==========================================
  // 印刷用レポートビルダー
  // ==========================================
  function buildPrintReport() {
    const proj = getCurrentProject();
    if (!proj) return;
    const container = $('#printReport');
    const tasks = proj.tasks || [];
    const imps = proj.improvements || [];
    const measures = proj.measurements || [];
    const companyName = proj.company || '';
    const companyInfo = (appData.companies || {})[companyName] || {};
    const todayStr = new Date().toLocaleDateString('ja-JP');

    let html = '';
    html += prCoverPage(proj, companyName, companyInfo, todayStr);
    html += prTOCPage();
    html += prExecutiveSummary(tasks, imps, measures);
    html += prStep1Page(tasks);
    html += prStep2Page(tasks);
    html += prStep3Page(tasks);
    html += prStep4Page(tasks, imps);
    html += prStep5Page(tasks, imps, measures);
    html += prAppendixPage(tasks);
    container.innerHTML = html;

    // Chart.js画像生成
    prGenerateChartImages(tasks);
  }

  function prCoverPage(proj, company, companyInfo, todayStr) {
    return `<div class="pr-page pr-cover">
      <div class="pr-cover-logo">BPI<span>Navi</span></div>
      <div class="pr-cover-title">業務プロセス改善 報告書</div>
      <div class="pr-cover-divider"></div>
      <div class="pr-cover-project">${escapeHtml(proj.name)}</div>
      <div class="pr-cover-company">${escapeHtml(company)}</div>
      ${companyInfo.description ? `<div class="pr-cover-company-desc">${escapeHtml(companyInfo.description)}</div>` : ''}
      <div class="pr-cover-date">作成日: ${todayStr}</div>
    </div>`;
  }

  function prTOCPage() {
    const items = [
      { num: '1', title: 'エグゼクティブサマリ' },
      { num: '2', title: 'Step 1: 業務プロセスの洗い出し' },
      { num: '3', title: 'Step 2: 問題点の可視化と分析' },
      { num: '4', title: 'Step 3: ECRSによる改善対象の選定' },
      { num: '5', title: 'Step 4: 具体的な改善策の検討' },
      { num: '6', title: 'Step 5: 効果検証と振り返り' },
      { num: null, title: '付録: 業務プロセス詳細一覧' }
    ];
    return `<div class="pr-page">
      <div class="pr-toc-title">目 次</div>
      <div class="pr-toc">
        ${items.map(i => `<div class="pr-toc-item">
          <span>${i.num ? `<span class="pr-toc-num">${i.num}</span>` : '<span class="pr-toc-num-text">付録</span>'}${escapeHtml(i.title)}</span>
        </div>`).join('')}
      </div>
    </div>`;
  }

  function prExecutiveSummary(tasks, imps, measures) {
    const totalTasks = tasks.length;
    const improvedCount = imps.length;
    const doneCount = imps.filter(i => i.status === 'done').length;
    let totalTimeSaved = 0;
    measures.forEach(m => {
      const task = tasks.find(t => t.id === m.taskId);
      if (task) totalTimeSaved += (parseFloat(task.timeRequired) || 0) - (parseFloat(m.afterTime) || 0);
    });
    const problemTasks = tasks.filter(t => t.problems && t.problems.length > 0);
    const ecrsCandidates = tasks.filter(t => t.ecrs && ['E','C','R','S'].some(k => t.ecrs[k] === 'yes' || t.ecrs[k] === 'maybe'));

    const ecrsBreakdown = { E: 0, C: 0, R: 0, S: 0 };
    imps.forEach(i => { if (ecrsBreakdown[i.ecrsKey] !== undefined) ecrsBreakdown[i.ecrsKey]++; });
    const ecrsColors = { E: '#8b5cf6', C: '#06b6d4', R: '#f43f5e', S: '#f97316' };

    return `<div class="pr-page">
      <div class="pr-section-header">
        <div class="pr-section-title"><span class="pr-step-badge">1</span> エグゼクティブサマリ</div>
        <div class="pr-section-subtitle">プロジェクト全体の成果概要</div>
      </div>
      <div class="pr-kpi-row">
        <div class="pr-kpi-box"><div class="pr-kpi-val">${totalTasks}</div><div class="pr-kpi-lbl">総業務プロセス数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${problemTasks.length}</div><div class="pr-kpi-lbl">問題検出業務数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${ecrsCandidates.length}</div><div class="pr-kpi-lbl">改善候補業務数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${improvedCount}</div><div class="pr-kpi-lbl">改善計画策定数</div></div>
      </div>
      <div class="pr-kpi-row">
        <div class="pr-kpi-box"><div class="pr-kpi-val">${doneCount}</div><div class="pr-kpi-lbl">改善完了件数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${measures.length}</div><div class="pr-kpi-lbl">効果測定済み件数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${Math.round(totalTimeSaved)}分</div><div class="pr-kpi-lbl">削減時間合計（1回あたり）</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${totalTasks > 0 ? Math.round((doneCount / Math.max(improvedCount, 1)) * 100) : 0}%</div><div class="pr-kpi-lbl">改善完了率</div></div>
      </div>
      <div class="pr-subsection">ECRS区分別 改善件数</div>
      <div class="pr-ecrs-summary">
        ${ECRS_DEFINITIONS.map(def => `<div class="pr-ecrs-box" style="border-top:3px solid ${ecrsColors[def.key]};">
          <div class="pr-ecrs-box-val" style="color:${ecrsColors[def.key]};">${ecrsBreakdown[def.key]}</div>
          <div class="pr-ecrs-box-lbl">${def.key} - ${def.name.split('（')[0]}</div>
        </div>`).join('')}
      </div>
      ${measures.length > 0 ? `
        <div class="pr-subsection">効果測定ハイライト</div>
        <table class="pr-table">
          <thead><tr><th>業務内容</th><th>改善前</th><th>改善後</th><th>削減率</th></tr></thead>
          <tbody>${measures.map(m => {
            const t = tasks.find(t2 => t2.id === m.taskId);
            if (!t) return '';
            const b = parseFloat(t.timeRequired) || 0;
            const a = parseFloat(m.afterTime) || 0;
            const pct = b > 0 ? Math.round(((b - a) / b) * 100) : 0;
            return `<tr><td>${escapeHtml(t.content)}</td><td>${b}${escapeHtml(t.timeUnit||'分')}</td><td>${a}${escapeHtml(t.timeUnit||'分')}</td><td class="${pct >= 0 ? 'pr-positive' : 'pr-negative'}">${pct >= 0 ? '-' : '+'}${Math.abs(pct)}%</td></tr>`;
          }).join('')}</tbody>
        </table>
      ` : ''}
      <div class="pr-footer">BPI Navigator - 業務プロセス改善報告書</div>
    </div>`;
  }

  function prStep1Page(tasks) {
    // プロセス区分別集計
    const catStats = {};
    tasks.forEach(t => {
      if (!catStats[t.category]) catStats[t.category] = { count: 0, totalTime: 0 };
      catStats[t.category].count++;
      catStats[t.category].totalTime += calcMonthlyTime(t);
    });
    const totalMonthly = tasks.reduce((s, t) => s + calcMonthlyTime(t), 0);

    return `<div class="pr-page">
      <div class="pr-section-header">
        <div class="pr-section-title"><span class="pr-step-badge">Step 1</span> 業務プロセスの洗い出し</div>
        <div class="pr-section-subtitle">対象業務の一覧と時間分析</div>
      </div>
      <div class="pr-kpi-row">
        <div class="pr-kpi-box"><div class="pr-kpi-val">${tasks.length}</div><div class="pr-kpi-lbl">登録業務数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${Object.keys(catStats).length}</div><div class="pr-kpi-lbl">プロセス区分数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${Math.round(totalMonthly)}</div><div class="pr-kpi-lbl">月間合計（分）</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${(totalMonthly / 60).toFixed(1)}</div><div class="pr-kpi-lbl">月間合計（時間）</div></div>
      </div>
      <div class="pr-subsection">プロセス区分別集計</div>
      <table class="pr-table">
        <thead><tr><th>プロセス区分</th><th>業務数</th><th>月間合計時間（分）</th><th>構成比</th></tr></thead>
        <tbody>${Object.entries(catStats).sort((a, b) => b[1].totalTime - a[1].totalTime).map(([cat, s]) => {
          const pct = totalMonthly > 0 ? Math.round((s.totalTime / totalMonthly) * 100) : 0;
          return `<tr><td>${escapeHtml(cat)}</td><td style="text-align:center">${s.count}</td><td style="text-align:right">${Math.round(s.totalTime)}</td><td style="text-align:right">${pct}%</td></tr>`;
        }).join('')}</tbody>
      </table>
      <div class="pr-subsection">業務一覧</div>
      <table class="pr-table pr-table-compact">
        <thead><tr><th>No.</th><th>コード</th><th>区分</th><th>作業内容</th><th>担当者</th><th>所要時間</th><th>頻度</th><th>月間（分）</th></tr></thead>
        <tbody>${tasks.map((t, i) => `<tr>
          <td>${i + 1}</td><td>${escapeHtml(t.code || '')}</td><td>${escapeHtml(t.category || '')}</td>
          <td>${escapeHtml(t.content || '')}</td><td>${escapeHtml(t.person || '')}</td>
          <td style="text-align:right">${t.timeRequired || ''}${escapeHtml(t.timeUnit || '分')}</td>
          <td>${escapeHtml(t.freqType || '')} ${t.freqCount || ''}回</td>
          <td style="text-align:right">${Math.round(calcMonthlyTime(t))}</td>
        </tr>`).join('')}</tbody>
      </table>
      <div class="pr-footer">BPI Navigator - 業務プロセス改善報告書</div>
    </div>`;
  }

  function prStep2Page(tasks) {
    // 問題カテゴリ集計
    const problemCounts = {};
    PROBLEM_CATEGORIES.forEach(pc => { problemCounts[pc.id] = 0; });
    tasks.forEach(t => (t.problems || []).forEach(p => { if (problemCounts[p] !== undefined) problemCounts[p]++; }));
    const sortedProblems = Object.entries(problemCounts).sort((a, b) => b[1] - a[1]);
    const maxProblem = sortedProblems.length > 0 ? sortedProblems[0][1] : 1;

    // スコアランキング（上位10件）
    const scored = tasks.filter(t => t.scores).sort((a, b) => calcTotalScore(b.scores) - calcTotalScore(a.scores)).slice(0, 10);

    // ヒートマップ
    const categories = [...new Set(tasks.map(t => t.category).filter(Boolean))];
    const matrix = {};
    categories.forEach(cat => { matrix[cat] = {}; PROBLEM_CATEGORIES.forEach(pc => { matrix[cat][pc.id] = 0; }); });
    tasks.forEach(t => { (t.problems || []).forEach(p => { if (matrix[t.category]) matrix[t.category][p]++; }); });
    const maxVal = Math.max(1, ...Object.values(matrix).flatMap(row => Object.values(row)));

    return `<div class="pr-page">
      <div class="pr-section-header">
        <div class="pr-section-title"><span class="pr-step-badge">Step 2</span> 問題点の可視化と分析</div>
        <div class="pr-section-subtitle">問題カテゴリ分布・優先度スコアリング・ヒートマップ</div>
      </div>
      <div class="pr-subsection">問題カテゴリ分布</div>
      <div class="pr-bar-chart">
        ${sortedProblems.map(([id, count]) => {
          const cat = PROBLEM_CATEGORIES.find(pc => pc.id === id);
          const pct = maxProblem > 0 ? Math.round((count / maxProblem) * 100) : 0;
          return `<div class="pr-bar-row">
            <div class="pr-bar-label">${cat ? escapeHtml(cat.name) : id}</div>
            <div class="pr-bar-track"><div class="pr-bar-fill" style="width:${pct}%"></div></div>
            <div class="pr-bar-value">${count}件</div>
          </div>`;
        }).join('')}
      </div>
      <div class="pr-subsection">パレート分析</div>
      <img id="prChartPareto" class="pr-chart-img" alt="パレート図" style="max-height:260px;">
      <div class="pr-subsection">優先度スコアランキング（上位10件）</div>
      <table class="pr-table">
        <thead><tr><th>順位</th><th>コード</th><th>作業内容</th><th>時間影響</th><th>品質影響</th><th>頻度</th><th>改善容易度</th><th>総合スコア</th></tr></thead>
        <tbody>${scored.map((t, i) => {
          const s = t.scores;
          return `<tr><td style="text-align:center;font-weight:700">${i + 1}</td><td>${escapeHtml(t.code || '')}</td><td>${escapeHtml(t.content || '')}</td>
            <td style="text-align:center">${s.timeImpact}</td><td style="text-align:center">${s.qualityImpact}</td>
            <td style="text-align:center">${s.frequency}</td><td style="text-align:center">${s.difficulty}</td>
            <td style="text-align:center;font-weight:700;color:#2563eb;">${calcTotalScore(s).toFixed(2)}</td></tr>`;
        }).join('')}</tbody>
      </table>
      <div class="pr-subsection">ヒートマップ（プロセス区分 × 問題カテゴリ）</div>
      <table class="pr-heatmap">
        <thead><tr><th></th>${PROBLEM_CATEGORIES.map(pc => `<th>${escapeHtml(pc.name)}</th>`).join('')}</tr></thead>
        <tbody>${categories.map(cat => `<tr><th style="text-align:left">${escapeHtml(cat)}</th>
          ${PROBLEM_CATEGORIES.map(pc => {
            const v = matrix[cat][pc.id];
            const lv = Math.min(5, Math.ceil((v / maxVal) * 5));
            return `<td class="pr-heat-${v > 0 ? lv : 0}">${v || ''}</td>`;
          }).join('')}
        </tr>`).join('')}</tbody>
      </table>
      <div class="pr-footer">BPI Navigator - 業務プロセス改善報告書</div>
    </div>`;
  }

  function prStep3Page(tasks) {
    const analyzed = tasks.filter(t => t.ecrs);
    const yesCount = { E: 0, C: 0, R: 0, S: 0 };
    const maybeCount = { E: 0, C: 0, R: 0, S: 0 };
    analyzed.forEach(t => {
      ['E', 'C', 'R', 'S'].forEach(k => {
        if (t.ecrs[k] === 'yes') yesCount[k]++;
        else if (t.ecrs[k] === 'maybe') maybeCount[k]++;
      });
    });
    const ecrsColors = { E: '#8b5cf6', C: '#06b6d4', R: '#f43f5e', S: '#f97316' };

    return `<div class="pr-page">
      <div class="pr-section-header">
        <div class="pr-section-title"><span class="pr-step-badge">Step 3</span> ECRSによる改善対象の選定</div>
        <div class="pr-section-subtitle">E(排除)・C(結合)・R(交換)・S(簡素化)の4視点で改善候補を特定</div>
      </div>
      <div class="pr-subsection">ECRS候補サマリ</div>
      <div class="pr-ecrs-summary">
        ${ECRS_DEFINITIONS.map(def => `<div class="pr-ecrs-box" style="border-top:3px solid ${ecrsColors[def.key]};">
          <div class="pr-ecrs-box-val" style="color:${ecrsColors[def.key]};">${yesCount[def.key]}<span style="font-size:12px;color:#94a3b8;"> + ${maybeCount[def.key]}</span></div>
          <div class="pr-ecrs-box-lbl">${def.key} - ${def.name.split('（')[0]}<br><span style="font-size:8px;">該当 + 検討余地</span></div>
        </div>`).join('')}
      </div>
      <div class="pr-subsection">ECRS分析結果一覧</div>
      <table class="pr-table">
        <thead><tr><th>コード</th><th>作業内容</th><th>スコア</th><th style="text-align:center">E 排除</th><th style="text-align:center">C 結合</th><th style="text-align:center">R 交換</th><th style="text-align:center">S 簡素化</th></tr></thead>
        <tbody>${analyzed.sort((a, b) => calcTotalScore(b.scores) - calcTotalScore(a.scores)).map(t => {
          const score = t.scores ? calcTotalScore(t.scores).toFixed(2) : '-';
          return `<tr><td>${escapeHtml(t.code || '')}</td><td>${escapeHtml(t.content || '')}</td><td style="text-align:center">${score}</td>
            ${['E', 'C', 'R', 'S'].map(k => {
              const v = t.ecrs[k];
              const cls = v === 'yes' ? 'pr-ecrs-cell-yes' : v === 'maybe' ? 'pr-ecrs-cell-maybe' : 'pr-ecrs-cell-no';
              const txt = v === 'yes' ? '◎' : v === 'maybe' ? '△' : '−';
              return `<td class="${cls}">${txt}</td>`;
            }).join('')}
          </tr>`;
        }).join('')}</tbody>
      </table>
      <div class="pr-footer">BPI Navigator - 業務プロセス改善報告書</div>
    </div>`;
  }

  function prStep4Page(tasks, imps) {
    const statusLabel = { pending: '未着手', progress: '進行中', done: '完了' };
    const statusCls = { pending: 'pr-status-pending', progress: 'pr-status-progress', done: 'pr-status-done' };
    const statusCounts = { pending: 0, progress: 0, done: 0 };
    imps.forEach(i => { if (statusCounts[i.status] !== undefined) statusCounts[i.status]++; });

    return `<div class="pr-page">
      <div class="pr-section-header">
        <div class="pr-section-title"><span class="pr-step-badge">Step 4</span> 具体的な改善策の検討</div>
        <div class="pr-section-subtitle">改善計画の策定と実施状況</div>
      </div>
      <div class="pr-kpi-row">
        <div class="pr-kpi-box"><div class="pr-kpi-val">${imps.length}</div><div class="pr-kpi-lbl">改善計画数</div></div>
        <div class="pr-kpi-box" style="border-top:2px solid #94a3b8"><div class="pr-kpi-val">${statusCounts.pending}</div><div class="pr-kpi-lbl">未着手</div></div>
        <div class="pr-kpi-box" style="border-top:2px solid #f59e0b"><div class="pr-kpi-val">${statusCounts.progress}</div><div class="pr-kpi-lbl">進行中</div></div>
        <div class="pr-kpi-box" style="border-top:2px solid #22c55e"><div class="pr-kpi-val">${statusCounts.done}</div><div class="pr-kpi-lbl">完了</div></div>
      </div>
      <div class="pr-subsection">改善計画一覧</div>
      <table class="pr-table pr-table-compact">
        <thead><tr><th>対象業務</th><th>ECRS</th><th>改善内容</th><th>改善手段</th><th>担当</th><th>期間</th><th>ステータス</th></tr></thead>
        <tbody>${imps.map(imp => {
          const task = tasks.find(t => t.id === imp.taskId);
          return `<tr>
            <td>${task ? escapeHtml(task.content) : '?'}</td>
            <td><span class="pr-ecrs pr-ecrs-${imp.ecrsKey.toLowerCase()}">${imp.ecrsKey}</span></td>
            <td>${escapeHtml(imp.content || '')}</td>
            <td>${escapeHtml(imp.method || '')}</td>
            <td>${escapeHtml(imp.person || '')}</td>
            <td style="white-space:nowrap">${imp.startDate || ''} 〜 ${imp.endDate || ''}</td>
            <td><span class="pr-status ${statusCls[imp.status] || ''}">${statusLabel[imp.status] || imp.status}</span></td>
          </tr>`;
        }).join('')}</tbody>
      </table>
      ${imps.some(i => i.effectQuantity || i.effectQuality) ? `
        <div class="pr-subsection">期待効果</div>
        <table class="pr-table pr-table-compact">
          <thead><tr><th>対象業務</th><th>ECRS</th><th>定量的効果</th><th>定性的効果</th></tr></thead>
          <tbody>${imps.filter(i => i.effectQuantity || i.effectQuality).map(imp => {
            const task = tasks.find(t => t.id === imp.taskId);
            return `<tr><td>${task ? escapeHtml(task.content) : '?'}</td><td><span class="pr-ecrs pr-ecrs-${imp.ecrsKey.toLowerCase()}">${imp.ecrsKey}</span></td>
              <td>${escapeHtml(imp.effectQuantity || '')}</td><td>${escapeHtml(imp.effectQuality || '')}</td></tr>`;
          }).join('')}</tbody>
        </table>
      ` : ''}
      <div class="pr-footer">BPI Navigator - 業務プロセス改善報告書</div>
    </div>`;
  }

  function prStep5Page(tasks, imps, measures) {
    let totalTimeSaved = 0;
    const rows = measures.map(m => {
      const task = tasks.find(t => t.id === m.taskId);
      if (!task) return null;
      const before = parseFloat(task.timeRequired) || 0;
      const after = parseFloat(m.afterTime) || 0;
      const saved = before - after;
      totalTimeSaved += saved;
      const pct = before > 0 ? Math.round(((before - after) / before) * 100) : 0;
      return { task, before, after, saved, pct, measure: m };
    }).filter(Boolean);

    const doneImps = imps.filter(i => i.status === 'done').length;
    const progressImps = imps.filter(i => i.status === 'progress').length;

    return `<div class="pr-page">
      <div class="pr-section-header">
        <div class="pr-section-title"><span class="pr-step-badge">Step 5</span> 効果検証と振り返り</div>
        <div class="pr-section-subtitle">改善施策の効果測定と成果の定量評価</div>
      </div>
      <div class="pr-kpi-row">
        <div class="pr-kpi-box"><div class="pr-kpi-val">${measures.length}</div><div class="pr-kpi-lbl">効果測定済み件数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${Math.round(totalTimeSaved)}分</div><div class="pr-kpi-lbl">削減時間合計（1回あたり）</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${doneImps}</div><div class="pr-kpi-lbl">改善完了件数</div></div>
        <div class="pr-kpi-box"><div class="pr-kpi-val">${progressImps}</div><div class="pr-kpi-lbl">進行中件数</div></div>
      </div>
      ${rows.length > 0 ? `
        <div class="pr-subsection">改善前後の比較</div>
        <table class="pr-table">
          <thead><tr><th>業務内容</th><th style="text-align:right">改善前</th><th style="text-align:right">改善後</th><th style="text-align:right">削減時間</th><th style="text-align:right">削減率</th><th>測定日</th><th>コメント</th></tr></thead>
          <tbody>${rows.map(r => `<tr>
            <td>${escapeHtml(r.task.content)}</td>
            <td style="text-align:right">${r.before}${escapeHtml(r.task.timeUnit || '分')}</td>
            <td style="text-align:right">${r.after}${escapeHtml(r.task.timeUnit || '分')}</td>
            <td style="text-align:right" class="${r.saved >= 0 ? 'pr-positive' : 'pr-negative'}">${r.saved >= 0 ? '-' : '+'}${Math.abs(r.saved)}${escapeHtml(r.task.timeUnit || '分')}</td>
            <td style="text-align:right" class="${r.pct >= 0 ? 'pr-positive' : 'pr-negative'}">${r.pct >= 0 ? '-' : '+'}${Math.abs(r.pct)}%</td>
            <td>${r.measure.date || ''}</td>
            <td>${escapeHtml(r.measure.comment || '')}</td>
          </tr>`).join('')}
          <tr style="font-weight:700;border-top:2px solid #2563eb;">
            <td>合計</td><td></td><td></td>
            <td style="text-align:right" class="${totalTimeSaved >= 0 ? 'pr-positive' : 'pr-negative'}">${totalTimeSaved >= 0 ? '-' : '+'}${Math.abs(Math.round(totalTimeSaved))}分</td>
            <td colspan="3"></td>
          </tr></tbody>
        </table>
      ` : '<p style="color:#94a3b8;">効果測定データはまだありません。</p>'}
      <div class="pr-footer">BPI Navigator - 業務プロセス改善報告書</div>
    </div>`;
  }

  function prAppendixPage(tasks) {
    return `<div class="pr-page">
      <div class="pr-section-header">
        <div class="pr-section-title"><span class="pr-toc-num-text">付録</span> 業務プロセス詳細一覧</div>
        <div class="pr-section-subtitle">全${tasks.length}件の業務プロセスの詳細データ</div>
      </div>
      <table class="pr-table pr-table-compact">
        <thead><tr><th>コード</th><th>区分</th><th>作業内容</th><th>担当</th><th>対象</th><th>方法</th><th>時間</th><th>頻度</th><th>ツール</th><th>問題</th><th>スコア</th><th>ECRS</th></tr></thead>
        <tbody>${tasks.map(t => {
          const probNames = (t.problems || []).map(p => { const c = PROBLEM_CATEGORIES.find(pc => pc.id === p); return c ? c.name : p; }).join(', ');
          const score = t.scores ? calcTotalScore(t.scores).toFixed(1) : '-';
          const ecrsStr = t.ecrs ? ['E', 'C', 'R', 'S'].map(k => t.ecrs[k] === 'yes' ? k : t.ecrs[k] === 'maybe' ? k + '?' : '').filter(Boolean).join(',') : '-';
          return `<tr>
            <td>${escapeHtml(t.code || '')}</td><td>${escapeHtml(t.category || '')}</td><td>${escapeHtml(t.content || '')}</td>
            <td>${escapeHtml(t.person || '')}</td><td>${escapeHtml(t.target || '')}</td><td>${escapeHtml(t.method || '')}</td>
            <td style="text-align:right;white-space:nowrap">${t.timeRequired || ''}${escapeHtml(t.timeUnit || '分')}</td>
            <td style="white-space:nowrap">${escapeHtml(t.freqType || '')}${t.freqCount || ''}回</td>
            <td>${escapeHtml(t.tools || '')}</td>
            <td style="font-size:8px">${escapeHtml(probNames) || '-'}</td>
            <td style="text-align:center">${score}</td>
            <td style="text-align:center;font-size:9px">${ecrsStr}</td>
          </tr>`;
        }).join('')}</tbody>
      </table>
      <div class="pr-footer">BPI Navigator - 業務プロセス改善報告書</div>
    </div>`;
  }

  function prGenerateChartImages(tasks) {
    const paretoTarget = $('#prChartPareto');
    if (!paretoTarget) return;

    const tempDiv = document.createElement('div');
    tempDiv.style.cssText = 'position:absolute;left:-9999px;top:-9999px;';
    document.body.appendChild(tempDiv);

    try {
      // パレート図
      const canvas = document.createElement('canvas');
      canvas.width = 600;
      canvas.height = 280;
      tempDiv.appendChild(canvas);

      const problemCounts = {};
      PROBLEM_CATEGORIES.forEach(pc => { problemCounts[pc.id] = 0; });
      tasks.forEach(t => (t.problems || []).forEach(p => { if (problemCounts[p] !== undefined) problemCounts[p]++; }));
      const sorted = Object.entries(problemCounts).sort((a, b) => b[1] - a[1]);
      const labels = sorted.map(([id]) => { const c = PROBLEM_CATEGORIES.find(pc => pc.id === id); return c ? c.name : id; });
      const values = sorted.map(([, v]) => v);
      const total = values.reduce((s, v) => s + v, 0) || 1;
      let cum = 0;
      const cumPct = values.map(v => { cum += v; return Math.round((cum / total) * 100); });

      const chart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
          labels,
          datasets: [
            { label: '件数', data: values, backgroundColor: 'rgba(37,99,235,0.6)', yAxisID: 'y', order: 2 },
            { label: '累積%', data: cumPct, type: 'line', borderColor: '#DC2626', backgroundColor: '#DC2626', yAxisID: 'y1', pointRadius: 3, borderWidth: 2, order: 1 }
          ]
        },
        options: {
          responsive: false, animation: false,
          plugins: { legend: { position: 'bottom', labels: { font: { size: 10 } } } },
          scales: {
            y: { beginAtZero: true, position: 'left', title: { display: true, text: '件数', font: { size: 10 } }, ticks: { font: { size: 9 } } },
            y1: { beginAtZero: true, max: 100, position: 'right', title: { display: true, text: '累積%', font: { size: 10 } }, grid: { drawOnChartArea: false }, ticks: { font: { size: 9 } } },
            x: { ticks: { font: { size: 9 }, maxRotation: 45 } }
          }
        }
      });
      paretoTarget.src = chart.toBase64Image();
      chart.destroy();
    } catch (e) {
      console.warn('Chart image generation failed:', e);
    } finally {
      document.body.removeChild(tempDiv);
    }
  }

  // ==========================================
  // Excelエクスポート
  // ==========================================
  function exportExcel() {
    if (typeof XLSX === 'undefined') {
      showToast('Excelライブラリが読み込まれていません', 'error');
      return;
    }
    const proj = getCurrentProject();
    if (!proj) { showToast('プロジェクトが選択されていません', 'error'); return; }

    const tasks = proj.tasks || [];
    const imps = proj.improvements || [];
    const measures = proj.measurements || [];
    const wb = XLSX.utils.book_new();

    xlSummarySheet(wb, proj, tasks, imps, measures);
    xlTaskListSheet(wb, tasks);
    xlProblemSheet(wb, tasks);
    xlECRSSheet(wb, tasks);
    xlImprovementSheet(wb, tasks, imps);
    xlMeasurementSheet(wb, tasks, imps, measures);

    const filename = `BPI_${sanitizeFilename(proj.company || 'export')}_${sanitizeFilename(proj.name || 'project')}_${today()}.xlsx`;
    XLSX.writeFile(wb, filename);
    showToast('Excelファイルをエクスポートしました', 'success');
  }

  function xlSummarySheet(wb, proj, tasks, imps, measures) {
    const doneCount = imps.filter(i => i.status === 'done').length;
    let totalTimeSaved = 0;
    measures.forEach(m => {
      const t = tasks.find(t2 => t2.id === m.taskId);
      if (t) totalTimeSaved += (parseFloat(t.timeRequired) || 0) - (parseFloat(m.afterTime) || 0);
    });
    const companyInfo = (appData.companies || {})[proj.company] || {};

    const data = [
      ['業務プロセス改善 サマリレポート'],
      [],
      ['プロジェクト名', proj.name || ''],
      ['企業名', proj.company || ''],
      ['企業概要', companyInfo.description || ''],
      ['作成日', today()],
      [],
      ['■ KPI', '値'],
      ['総業務プロセス数', tasks.length],
      ['改善対象件数', imps.length],
      ['改善完了件数', doneCount],
      ['効果測定済み件数', measures.length],
      ['削減時間合計（分）', Math.round(totalTimeSaved)],
      ['改善完了率（%）', imps.length > 0 ? Math.round((doneCount / imps.length) * 100) : 0],
      [],
      ['■ ECRS区分別', '件数'],
      ['E（排除）', imps.filter(i => i.ecrsKey === 'E').length],
      ['C（結合）', imps.filter(i => i.ecrsKey === 'C').length],
      ['R（交換）', imps.filter(i => i.ecrsKey === 'R').length],
      ['S（簡素化）', imps.filter(i => i.ecrsKey === 'S').length]
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 28 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, ws, 'サマリ');
  }

  function xlTaskListSheet(wb, tasks) {
    const header = ['コード', 'プロセス区分', '作業内容', '担当者', '対象', '方法', '所要時間', '単位', '頻度種別', '頻度回数', '月間時間（分）', 'ツール', 'ムリ・ムダ・ムラ'];
    const rows = tasks.map(t => [
      t.code || '', t.category || '', t.content || '', t.person || '', t.target || '', t.method || '',
      t.timeRequired || '', t.timeUnit || '分', t.freqType || '', t.freqCount || '',
      Math.round(calcMonthlyTime(t)), t.tools || '', (t.mmm || []).join('・') || t.notes || ''
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws['!cols'] = [{ wch: 10 }, { wch: 12 }, { wch: 30 }, { wch: 10 }, { wch: 15 }, { wch: 20 }, { wch: 8 }, { wch: 5 }, { wch: 8 }, { wch: 6 }, { wch: 10 }, { wch: 15 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, ws, '業務一覧');
  }

  function xlProblemSheet(wb, tasks) {
    // 問題カテゴリ集計
    const data = [['■ 問題カテゴリ集計'], ['カテゴリ', '件数']];
    const counts = {};
    PROBLEM_CATEGORIES.forEach(pc => { counts[pc.id] = 0; });
    tasks.forEach(t => (t.problems || []).forEach(p => { if (counts[p] !== undefined) counts[p]++; }));
    Object.entries(counts).sort((a, b) => b[1] - a[1]).forEach(([id, cnt]) => {
      const cat = PROBLEM_CATEGORIES.find(pc => pc.id === id);
      data.push([cat ? cat.name : id, cnt]);
    });

    // スコアランキング
    data.push([], ['■ 優先度スコアランキング'], ['順位', 'コード', '作業内容', '時間影響', '品質影響', '頻度', '改善容易度', '総合スコア']);
    tasks.filter(t => t.scores).sort((a, b) => calcTotalScore(b.scores) - calcTotalScore(a.scores)).forEach((t, i) => {
      const s = t.scores;
      data.push([i + 1, t.code || '', t.content || '', s.timeImpact, s.qualityImpact, s.frequency, s.difficulty, parseFloat(calcTotalScore(s).toFixed(2))]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 8 }, { wch: 12 }, { wch: 30 }, { wch: 8 }, { wch: 8 }, { wch: 6 }, { wch: 10 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, '問題分析');
  }

  function xlECRSSheet(wb, tasks) {
    const header = ['コード', '作業内容', '優先度スコア', 'E（排除）', 'C（結合）', 'R（交換）', 'S（簡素化）'];
    const label = { yes: '◎ はい', maybe: '△ 検討', no: '− いいえ' };
    const rows = tasks.filter(t => t.ecrs).sort((a, b) => calcTotalScore(b.scores) - calcTotalScore(a.scores)).map(t => [
      t.code || '', t.content || '', t.scores ? parseFloat(calcTotalScore(t.scores).toFixed(2)) : '',
      label[t.ecrs.E] || '-', label[t.ecrs.C] || '-', label[t.ecrs.R] || '-', label[t.ecrs.S] || '-'
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws['!cols'] = [{ wch: 10 }, { wch: 30 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, ws, 'ECRS分析');
  }

  function xlImprovementSheet(wb, tasks, imps) {
    const header = ['対象業務', 'ECRS区分', '改善内容', '改善手段', '担当者', 'リソース', '開始日', '完了日', '期待効果（定量）', '期待効果（定性）', 'ステータス'];
    const statusLabel = { pending: '未着手', progress: '進行中', done: '完了' };
    const rows = imps.map(imp => {
      const task = tasks.find(t => t.id === imp.taskId);
      return [task ? task.content : '?', imp.ecrsKey, imp.content || '', imp.method || '', imp.person || '',
        imp.resources || '', imp.startDate || '', imp.endDate || '', imp.effectQuantity || '', imp.effectQuality || '',
        statusLabel[imp.status] || imp.status];
    });
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws['!cols'] = [{ wch: 25 }, { wch: 8 }, { wch: 25 }, { wch: 20 }, { wch: 10 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 20 }, { wch: 20 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, '改善計画');
  }

  function xlMeasurementSheet(wb, tasks, imps, measures) {
    const header = ['業務内容', '改善前時間', '単位', '改善後時間', '削減時間', '削減率（%）', '達成率', '測定日', 'コメント'];
    const rows = measures.map(m => {
      const task = tasks.find(t => t.id === m.taskId);
      if (!task) return ['?', '', '', '', '', '', '', '', ''];
      const before = parseFloat(task.timeRequired) || 0;
      const after = parseFloat(m.afterTime) || 0;
      const saved = before - after;
      const pct = before > 0 ? Math.round(((before - after) / before) * 100) : 0;
      return [task.content || '', before, task.timeUnit || '分', after, saved, pct, m.achievement || '', m.date || '', m.comment || ''];
    });
    // 合計行
    if (rows.length > 0) {
      const totalSaved = rows.reduce((s, r) => s + (parseFloat(r[4]) || 0), 0);
      rows.push(['【合計】', '', '', '', Math.round(totalSaved), '', '', '', '']);
    }
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws['!cols'] = [{ wch: 30 }, { wch: 10 }, { wch: 5 }, { wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 8 }, { wch: 12 }, { wch: 25 }];
    XLSX.utils.book_append_sheet(wb, ws, '効果測定');
  }

  // ==========================================
  // 設定・マスタ管理
  // ==========================================
  function renderSettings() {
    renderProcessCategoryMaster();
    renderProblemCategoryMaster();
    renderTemplates();
    renderAISettings();
  }

  function renderProcessCategoryMaster() {
    const container = $('#processCategoryMaster');
    container.innerHTML = '<div class="master-list">' +
      appData.processCategories.map(c => `
        <div class="master-item">
          ${escapeHtml(c)}
          <button class="remove-btn" data-cat="${escapeHtml(c)}">&times;</button>
        </div>
      `).join('') + '</div>';

    container.querySelectorAll('.remove-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        appData.processCategories = appData.processCategories.filter(c => c !== btn.dataset.cat);
        saveData(appData);
        renderSettings();
        showToast('プロセス区分を削除しました');
      });
    });
  }

  function renderProblemCategoryMaster() {
    const container = $('#problemCategoryMaster');
    container.innerHTML = '<div class="master-list">' +
      PROBLEM_CATEGORIES.map(c =>
        `<div class="master-item">${escapeHtml(c.name)}</div>`
      ).join('') + '</div>';
  }

  function renderTemplates() {
    const container = $('#templateList');
    let html = '';
    Object.entries(IMPROVEMENT_TEMPLATES).forEach(([key, templates]) => {
      const cat = PROBLEM_CATEGORIES.find(pc => pc.id === key);
      if (!cat) return;
      html += `<div class="template-item"><strong>${escapeHtml(cat.name)}</strong><br>`;
      html += templates.map(t => `・${escapeHtml(t)}`).join('<br>');
      html += '</div>';
    });
    container.innerHTML = html;
  }

  // ==========================================
  // AI設定 レンダリング
  // ==========================================
  function renderAISettings() {
    const cfg = AIAssist.loadConfig();
    const provSel = $('#aiProvider');
    const keySel = $('#aiApiKey');
    const modelSel = $('#aiModel');
    if (!provSel) return;

    provSel.value = cfg.provider || 'claude';
    keySel.value = cfg.apiKey || '';
    updateModelOptions(cfg.provider, cfg.model);

    provSel.onchange = () => {
      updateModelOptions(provSel.value, '');
    };

    $('#btnToggleApiKey').onclick = () => {
      const inp = $('#aiApiKey');
      if (inp.type === 'password') { inp.type = 'text'; $('#btnToggleApiKey').textContent = '隠す'; }
      else { inp.type = 'password'; $('#btnToggleApiKey').textContent = '表示'; }
    };

    $('#btnSaveAIConfig').onclick = () => {
      const newCfg = {
        provider: provSel.value,
        apiKey: keySel.value.trim(),
        model: modelSel.value
      };
      AIAssist.saveConfig(newCfg);
      showToast('AI設定を保存しました', 'success');
    };

    $('#btnTestAI').onclick = async () => {
      const resultEl = $('#aiTestResult');
      // 一時的に入力値で保存してからテスト
      const tmpCfg = {
        provider: provSel.value,
        apiKey: keySel.value.trim(),
        model: modelSel.value
      };
      AIAssist.saveConfig(tmpCfg);
      resultEl.innerHTML = '<span style="color:var(--text-secondary)">接続テスト中...</span>';
      try {
        const reply = await AIAssist.testConnection();
        resultEl.innerHTML = `<span style="color:var(--success)">✅ 接続成功: ${escapeHtml(reply.slice(0, 50))}</span>`;
      } catch (e) {
        resultEl.innerHTML = `<span style="color:var(--danger)">❌ エラー: ${escapeHtml(e.message)}</span>`;
      }
    };
  }

  function updateModelOptions(provider, selectedModel) {
    const modelSel = $('#aiModel');
    const prov = AIAssist.PROVIDERS[provider];
    if (!prov) return;
    modelSel.innerHTML = prov.models.map(m =>
      `<option value="${m.id}" ${m.id === selectedModel ? 'selected' : ''}>${m.name}</option>`
    ).join('');
    if (!selectedModel) modelSel.value = prov.defaultModel;
  }

  // ==========================================
  // AIアシスト モーダル制御
  // ==========================================
  let aiCurrentStep = 1;
  let aiLastResult = null;

  const STEP_TITLES = {
    1: 'Step 1: 業務の洗い出し',
    2: 'Step 2: 問題分析・スコアリング',
    3: 'Step 3: ECRS分析',
    4: 'Step 4: 改善策の提案',
    5: 'Step 5: 効果検証の提案'
  };

  function openAIAssist(step) {
    const cfg = AIAssist.loadConfig();
    if (!cfg.apiKey) {
      showToast('AIのAPIキーが未設定です。設定画面で入力してください。', 'error');
      return;
    }
    aiCurrentStep = step;
    aiLastResult = null;
    $('#aiAssistTitle').textContent = `🤖 AIアシスト — ${STEP_TITLES[step] || ''}`;
    $('#aiExtraPrompt').value = '';
    $('#aiResultArea').innerHTML = '';
    $('#btnAIApply').style.display = 'none';
    $('#btnAIGenerate').style.display = '';

    // コンテキスト情報を表示
    const proj = getCurrentProject();
    const cfg2 = AIAssist.loadConfig();
    const provName = AIAssist.PROVIDERS[cfg2.provider]?.name || cfg2.provider;
    const modelName = AIAssist.getEffectiveModel(cfg2);
    let ctx = `<strong>企業:</strong> ${escapeHtml(proj?.company || '未設定')}`;
    ctx += ` ／ <strong>プロジェクト:</strong> ${escapeHtml(proj?.name || '')}`;
    ctx += ` ／ <strong>AI:</strong> ${escapeHtml(provName)} (${escapeHtml(modelName)})`;
    if (step === 1) ctx += `<br>登録済み業務: ${(proj?.tasks || []).length}件`;
    if (step === 2) ctx += `<br>対象業務: ${(proj?.tasks || []).length}件`;
    if (step === 3) {
      const scored = (proj?.tasks || []).filter(t => t.scores && t.problems?.length > 0).length;
      ctx += `<br>スコアリング済み業務: ${scored}件`;
    }
    if (step === 4) {
      const ecrsed = (proj?.tasks || []).filter(t => t.ecrs).length;
      ctx += `<br>ECRS分析済み業務: ${ecrsed}件`;
    }
    if (step === 5) ctx += `<br>改善計画: ${(proj?.improvements || []).length}件`;
    $('#aiContextInfo').innerHTML = ctx;

    openModal('modalAIAssist');
  }

  async function generateAISuggestion() {
    const proj = getCurrentProject();
    if (!proj) return;
    const resultArea = $('#aiResultArea');
    const extra = $('#aiExtraPrompt').value.trim();

    resultArea.innerHTML = '<div class="ai-loading"><div class="ai-spinner"></div> AIが分析中です... しばらくお待ちください</div>';
    $('#btnAIGenerate').disabled = true;
    $('#btnAIApply').style.display = 'none';

    try {
      const builders = {
        1: AIAssist.buildStep1Prompt,
        2: AIAssist.buildStep2Prompt,
        3: AIAssist.buildStep3Prompt,
        4: AIAssist.buildStep4Prompt,
        5: AIAssist.buildStep5Prompt
      };
      const builder = builders[aiCurrentStep];
      if (!builder) throw new Error('不明なStep');
      // 企業概要をプロジェクト情報に付与してAIに渡す
      const companyInfo = (appData.companies || {})[proj.company];
      const projForAI = { ...proj, companyDescription: companyInfo?.description || '' };
      const { system, user } = builder(projForAI);
      const rawResponse = await AIAssist.callAI(system, user, extra);
      const parsed = AIAssist.parseResponse(rawResponse);

      if (parsed.ok && Array.isArray(parsed.data)) {
        aiLastResult = parsed.data;
        renderAIResult(parsed.data);
        $('#btnAIApply').style.display = '';
      } else if (parsed.ok && typeof parsed.data === 'object') {
        aiLastResult = Array.isArray(parsed.data) ? parsed.data : [parsed.data];
        renderAIResult(aiLastResult);
        $('#btnAIApply').style.display = '';
      } else {
        resultArea.innerHTML = `<div class="ai-raw-result"><strong>AIの回答（テキスト）:</strong><pre>${escapeHtml(parsed.raw || rawResponse)}</pre></div>`;
      }
    } catch (e) {
      resultArea.innerHTML = `<div class="ai-error">❌ エラー: ${escapeHtml(e.message)}</div>`;
    } finally {
      $('#btnAIGenerate').disabled = false;
    }
  }

  function renderAIResult(data) {
    const area = $('#aiResultArea');
    if (!data || data.length === 0) {
      area.innerHTML = '<p>提案がありませんでした。</p>';
      return;
    }

    let html = `<div class="ai-result-header">提案: <strong>${data.length}件</strong>
      <label style="margin-left:12px;"><input type="checkbox" id="aiSelectAll" checked onchange="window.BPI.toggleAISelectAll(this.checked)"> すべて選択</label></div>`;
    html += '<div class="ai-result-list">';

    if (aiCurrentStep === 1) {
      data.forEach((item, i) => {
        html += `<div class="ai-result-card">
          <label><input type="checkbox" class="ai-check" data-idx="${i}" checked></label>
          <div class="ai-card-body">
            <strong>${escapeHtml(item.content || '')}</strong>
            <div class="ai-card-meta">担当: ${escapeHtml(item.person || '')} ／ 対象: ${escapeHtml(item.target || '')} ／ ${item.timeRequired || ''}${item.timeUnit || ''} ／ ${item.freqType || ''} ${item.freqCount || ''}回</div>
            <div class="ai-card-detail">${escapeHtml(item.method || '')}</div>
            ${item.notes ? `<div class="ai-card-note">📝 ${escapeHtml(item.notes)}</div>` : ''}
          </div></div>`;
      });
    } else if (aiCurrentStep === 2) {
      data.forEach((item, i) => {
        const probNames = (item.problems || []).map(p => {
          const cat = PROBLEM_CATEGORIES.find(c => c.id === p);
          return cat ? cat.name : p;
        }).join(', ');
        html += `<div class="ai-result-card">
          <label><input type="checkbox" class="ai-check" data-idx="${i}" checked></label>
          <div class="ai-card-body">
            <strong>${escapeHtml(item.code || '')}</strong>
            <div class="ai-card-meta">問題: ${escapeHtml(probNames)}</div>
            <div class="ai-card-meta">スコア: 時間${item.scores?.timeImpact || 0} / 品質${item.scores?.qualityImpact || 0} / 頻度${item.scores?.frequency || 0} / 難易度${item.scores?.difficulty || 0}</div>
            ${item.reason ? `<div class="ai-card-note">💡 ${escapeHtml(item.reason)}</div>` : ''}
          </div></div>`;
      });
    } else if (aiCurrentStep === 3) {
      data.forEach((item, i) => {
        const ecrs = item.ecrs || {};
        const labels = ['E', 'C', 'R', 'S'].map(k => `${k}:${ecrs[k] || 'no'}`).join(' / ');
        html += `<div class="ai-result-card">
          <label><input type="checkbox" class="ai-check" data-idx="${i}" checked></label>
          <div class="ai-card-body">
            <strong>${escapeHtml(item.code || '')}</strong>
            <div class="ai-card-meta">${escapeHtml(labels)}</div>
            ${item.reason ? `<div class="ai-card-note">💡 ${escapeHtml(item.reason)}</div>` : ''}
          </div></div>`;
      });
    } else if (aiCurrentStep === 4) {
      data.forEach((item, i) => {
        html += `<div class="ai-result-card">
          <label><input type="checkbox" class="ai-check" data-idx="${i}" checked></label>
          <div class="ai-card-body">
            <strong>[${escapeHtml(item.taskCode || '')}/${item.ecrsKey || ''}] ${escapeHtml(item.content || '')}</strong>
            <div class="ai-card-detail">${escapeHtml(item.method || '')}</div>
            <div class="ai-card-meta">担当: ${escapeHtml(item.person || '')} ／ リソース: ${escapeHtml(item.resources || '')}</div>
            <div class="ai-card-meta">期待効果: ${escapeHtml(item.effectQuantity || '')} ／ ${escapeHtml(item.effectQuality || '')}</div>
          </div></div>`;
      });
    } else if (aiCurrentStep === 5) {
      data.forEach((item, i) => {
        html += `<div class="ai-result-card">
          <label><input type="checkbox" class="ai-check" data-idx="${i}" checked></label>
          <div class="ai-card-body">
            <strong>[${escapeHtml(item.taskCode || '')}] KPI: ${escapeHtml(item.suggestedKPI || '')}</strong>
            <div class="ai-card-detail">測定方法: ${escapeHtml(item.measurementMethod || '')}</div>
            <div class="ai-card-meta">目標値: ${escapeHtml(item.expectedValue || '')}</div>
            ${item.checkpoints ? `<div class="ai-card-note">✅ ${escapeHtml(item.checkpoints)}</div>` : ''}
          </div></div>`;
      });
    }

    html += '</div>';
    area.innerHTML = html;
  }

  function toggleAISelectAll(checked) {
    $$('#aiResultArea .ai-check').forEach(cb => { cb.checked = checked; });
  }

  function getSelectedAIIndices() {
    const indices = [];
    $$('#aiResultArea .ai-check').forEach(cb => {
      if (cb.checked) indices.push(parseInt(cb.dataset.idx));
    });
    return indices;
  }

  // ==========================================
  // AI提案の適用
  // ==========================================
  function applyAISuggestion() {
    if (!aiLastResult) return;
    const proj = getCurrentProject();
    if (!proj) return;
    const selected = getSelectedAIIndices();
    if (selected.length === 0) {
      showToast('適用する提案を選択してください', 'error');
      return;
    }

    const items = selected.map(i => aiLastResult[i]).filter(Boolean);

    switch (aiCurrentStep) {
      case 1: applyAITasks(proj, items); break;
      case 2: applyAIScores(proj, items); break;
      case 3: applyAIEcrs(proj, items); break;
      case 4: applyAIImprovements(proj, items); break;
      case 5: showToast('Step 5の提案は参考情報です。各改善策の効果測定画面で入力してください。', 'success'); break;
    }

    closeModal('modalAIAssist');
  }

  function applyAITasks(proj, items) {
    if (!proj.tasks) proj.tasks = [];
    const cat = (proj.processes && proj.processes[0]) || 'その他';
    let added = 0;
    items.forEach(item => {
      if (!item.content) return;
      const code = generateTaskCode(cat);
      proj.tasks.push({
        id: genId(),
        category: cat,
        code,
        content: item.content,
        person: item.person || '',
        target: item.target || '',
        method: item.method || '',
        timeRequired: item.timeRequired || '',
        timeUnit: item.timeUnit || '分',
        freqType: item.freqType || '日次',
        freqCount: item.freqCount || '1',
        tools: item.tools || '',
        notes: item.notes || '',
        mmm: item.mmm || [],
        scores: null,
        problems: [],
        ecrs: null
      });
      added++;
    });
    saveData(appData);
    renderStep1();
    showToast(`${added}件の業務を追加しました`, 'success');
  }

  function applyAIScores(proj, items) {
    let updated = 0;
    items.forEach(item => {
      if (!item.code) return;
      const task = (proj.tasks || []).find(t => t.code === item.code);
      if (!task) return;
      if (item.problems && Array.isArray(item.problems)) task.problems = item.problems;
      if (item.scores) task.scores = item.scores;
      updated++;
    });
    saveData(appData);
    navigate('step2');
    showToast(`${updated}件のスコアリングを適用しました`, 'success');
  }

  function applyAIEcrs(proj, items) {
    let updated = 0;
    items.forEach(item => {
      if (!item.code) return;
      const task = (proj.tasks || []).find(t => t.code === item.code);
      if (!task) return;
      if (item.ecrs) task.ecrs = item.ecrs;
      updated++;
    });
    saveData(appData);
    navigate('step3');
    showToast(`${updated}件のECRS分析を適用しました`, 'success');
  }

  function applyAIImprovements(proj, items) {
    if (!proj.improvements) proj.improvements = [];
    let added = 0;
    items.forEach(item => {
      if (!item.taskCode || !item.content) return;
      const task = (proj.tasks || []).find(t => t.code === item.taskCode);
      if (!task) return;
      // 同じタスク+ECRSキーで既存チェック
      const exists = proj.improvements.find(
        imp => imp.taskId === task.id && imp.ecrsKey === (item.ecrsKey || 'S')
      );
      if (exists) return;
      proj.improvements.push({
        taskId: task.id,
        ecrsKey: item.ecrsKey || 'S',
        content: item.content,
        method: item.method || '',
        person: item.person || '',
        resources: item.resources || '',
        startDate: '',
        endDate: '',
        effectQuantity: item.effectQuantity || '',
        effectQuality: item.effectQuality || '',
        status: 'pending'
      });
      added++;
    });
    saveData(appData);
    navigate('step4');
    showToast(`${added}件の改善計画を追加しました`, 'success');
  }

  // ==========================================
  // 企業管理
  // ==========================================
  let editingCompanyName = null; // 編集中の企業名（null=新規）

  function openCompanyModal(companyName) {
    editingCompanyName = companyName || null;
    if (editingCompanyName) {
      $('#modalCompanyTitle').textContent = '企業情報を編集';
      $('#companyName').value = editingCompanyName;
      $('#companyName').readOnly = true;
      const info = appData.companies[editingCompanyName] || {};
      $('#companyDescription').value = info.description || '';
    } else {
      $('#modalCompanyTitle').textContent = '企業を登録';
      $('#companyName').value = '';
      $('#companyName').readOnly = false;
      $('#companyDescription').value = '';
    }
    openModal('modalCompany');
  }

  function saveCompanyInfo() {
    const name = $('#companyName').value.trim();
    if (!name) {
      showToast('企業名は必須です', 'error');
      return;
    }
    const description = $('#companyDescription').value.trim();

    if (!editingCompanyName && appData.companies[name]) {
      showToast('この企業名は既に登録されています', 'error');
      return;
    }

    appData.companies[name] = { description };
    saveData(appData);
    closeModal('modalCompany');
    showToast(editingCompanyName ? '企業情報を更新しました' : '企業を登録しました', 'success');

    // プロジェクト作成モーダルが開いている場合、企業セレクトを更新して新企業を選択
    if ($('#modalProject').classList.contains('active')) {
      populateCompanySelect(name);
    }

    renderDashboard();
  }

  // ==========================================
  // プロジェクト作成
  // ==========================================
  /** 企業selectの選択肢を動的生成 */
  function populateCompanySelect(presetCompany) {
    const sel = $('#projectCompany');
    // 既存の動的オプションを除去（最初の「選択してください」と最後の「新しい企業を登録」は残す）
    sel.innerHTML = '<option value="">-- 選択してください --</option>';
    const names = Object.keys(appData.companies || {}).sort();
    names.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      sel.appendChild(opt);
    });
    const newOpt = document.createElement('option');
    newOpt.value = '__new__';
    newOpt.textContent = '＋ 新しい企業を登録...';
    sel.appendChild(newOpt);

    if (presetCompany) sel.value = presetCompany;
  }

  function openNewProjectModal(presetCompany) {
    $('#projectName').value = '';
    $('#projectMemo').value = '';
    $('#modalProjectTitle').textContent = '新規プロジェクト作成';

    // 企業セレクト
    populateCompanySelect(presetCompany || '');

    // プロセスチェックボックス
    const pg = $('#projectProcesses');
    pg.innerHTML = appData.processCategories.map(c =>
      `<label><input type="checkbox" value="${escapeHtml(c)}"> ${escapeHtml(c)}</label>`
    ).join('');

    openModal('modalProject');
  }

  function saveProject() {
    const name = $('#projectName').value.trim();
    const company = $('#projectCompany').value;
    if (!name || !company || company === '__new__') {
      showToast('プロジェクト名と企業名は必須です', 'error');
      return;
    }

    // 企業がcompaniesに未登録なら自動登録（セレクトから選ばれるので通常はありえないが念のため）
    if (!appData.companies[company]) {
      appData.companies[company] = { description: '' };
    }

    const selectedProcesses = [];
    $$('#projectProcesses input:checked').forEach(cb => selectedProcesses.push(cb.value));

    const project = {
      id: genId(),
      name,
      company,
      memo: $('#projectMemo').value.trim(),
      processes: selectedProcesses,
      currentStep: 1,
      tasks: [],
      improvements: [],
      measurements: [],
      createdAt: new Date().toISOString()
    };

    appData.projects.push(project);
    saveData(appData);
    closeModal('modalProject');
    selectProject(project.id);
    showToast('プロジェクトを作成しました', 'success');
  }

  // ==========================================
  // データエクスポート/インポート
  // ==========================================
  function exportAll() {
    const json = JSON.stringify(appData, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `bpi_navi_export_${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('データをエクスポートしました', 'success');
  }

  function importAll() {
    const fileInput = $('#importFileInput');
    fileInput.click();
    fileInput.onchange = () => {
      const file = fileInput.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const imported = JSON.parse(e.target.result);
          if (imported.projects) {
            appData = migrateData(imported);
            saveData(appData);
            navigate('dashboard');
            showToast('データをインポートしました', 'success');
          } else {
            showToast('無効なデータ形式です', 'error');
          }
        } catch (err) {
          showToast('インポートに失敗しました', 'error');
        }
      };
      reader.readAsText(file);
      fileInput.value = '';
    };
  }

  // ==========================================
  // イベントリスナー
  // ==========================================
  function initEventListeners() {
    // ナビゲーション
    $$('.nav-item').forEach(item => {
      item.addEventListener('click', () => {
        const view = item.dataset.view;
        if (view) {
          if (view !== 'dashboard' && view !== 'settings' && !getCurrentProject()) {
            showToast('先にプロジェクトを選択してください', 'error');
            return;
          }
          navigate(view);
        }
      });
    });

    // モバイルメニュー
    $('#menuToggle').addEventListener('click', () => {
      $('#sidebar').classList.toggle('open');
    });

    // モーダル閉じるボタン
    $$('[data-close]').forEach(btn => {
      btn.addEventListener('click', () => closeModal(btn.dataset.close));
    });

    // モーダルオーバーレイクリックで閉じる
    $$('.modal-overlay').forEach(overlay => {
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) overlay.classList.remove('active');
      });
    });

    // ドロップダウンメニュー制御
    setupDropdowns();

    // ファイル保存・読込
    $('#btnFileSaveAll').addEventListener('click', () => { closeAllDropdowns(); fileSaveAll(); });
    $('#btnFileSaveCompany').addEventListener('click', () => { closeAllDropdowns(); fileSaveCompany(); });
    $('#btnFileSaveProj').addEventListener('click', () => { closeAllDropdowns(); fileSaveProject(); });
    $('#btnFileLoadAll').addEventListener('click', () => { closeAllDropdowns(); fileLoadAll(); });
    $('#btnFileLoadMerge').addEventListener('click', () => { closeAllDropdowns(); fileLoadMerge(); });

    // 企業登録
    $('#btnAddCompany').addEventListener('click', () => openCompanyModal(null));
    $('#btnSaveCompany').addEventListener('click', saveCompanyInfo);

    // プロジェクト作成
    $('#btnNewProject').addEventListener('click', () => openNewProjectModal());
    $('#btnSaveProject').addEventListener('click', saveProject);

    // 企業セレクトで「新しい企業を登録」選択時
    $('#projectCompany').addEventListener('change', (e) => {
      if (e.target.value === '__new__') {
        // 企業登録モーダルを開く（プロジェクトモーダルはそのまま）
        openCompanyModal(null);
        // 登録後にプロジェクトモーダルのセレクトを更新するため
        // saveCompanyInfoの後にpopulateCompanySelectが呼ばれる仕組みを後で対応
      }
    });

    // Step1
    $('#btnAddTask').addEventListener('click', () => openTaskModal(null));
    $('#btnSaveTask').addEventListener('click', saveTask);
    $('#btnImportCSV').addEventListener('click', importCSV);
    $('#btnToggleMap').addEventListener('click', () => {
      const area = $('#processMapArea');
      if (area.style.display === 'none') {
        area.style.display = '';
        renderProcessMap();
      } else {
        area.style.display = 'none';
      }
    });

    // Step間ナビゲーション
    $('#btnGoStep2').addEventListener('click', () => {
      const proj = getCurrentProject();
      if (proj) {
        if (proj.currentStep < 2) proj.currentStep = 2;
        saveData(appData);
      }
      navigate('step2');
    });
    $('#btnBackStep1').addEventListener('click', () => navigate('step1'));
    $('#btnGoStep3').addEventListener('click', () => {
      const proj = getCurrentProject();
      if (proj) {
        if (proj.currentStep < 3) proj.currentStep = 3;
        saveData(appData);
      }
      navigate('step3');
    });
    $('#btnBackStep2').addEventListener('click', () => navigate('step2'));
    $('#btnGoStep4').addEventListener('click', () => {
      const proj = getCurrentProject();
      if (proj) {
        if (proj.currentStep < 4) proj.currentStep = 4;
        saveData(appData);
      }
      navigate('step4');
    });
    $('#btnBackStep3').addEventListener('click', () => navigate('step3'));
    $('#btnGoStep5').addEventListener('click', () => {
      const proj = getCurrentProject();
      if (proj) {
        if (proj.currentStep < 5) proj.currentStep = 5;
        saveData(appData);
      }
      navigate('step5');
    });
    $('#btnBackStep4').addEventListener('click', () => navigate('step4'));
    $('#btnNewCycle').addEventListener('click', () => {
      const proj = getCurrentProject();
      if (proj) {
        proj.currentStep = 1;
        saveData(appData);
        showToast('新しい改善サイクルを開始します', 'success');
        navigate('step1');
      }
    });

    // Step2 サブビュー切り替え
    $('#btnStep2Scoring').addEventListener('click', () => {
      switchSubView('step2', 'scoring');
      renderScoringList(getCurrentProject());
    });
    $('#btnStep2Dashboard').addEventListener('click', () => {
      switchSubView('step2', 'analysis');
      renderAnalysisDashboard();
    });

    // Step3 サブビュー切り替え
    $('#btnStep3Wizard').addEventListener('click', () => {
      switchSubView('step3', 'wizard');
      renderECRSWizard(getCurrentProject());
    });
    $('#btnStep3List').addEventListener('click', () => {
      switchSubView('step3', 'candidates');
      renderCandidatesList();
    });

    // Step4 サブビュー切り替え
    $('#btnStep4Plans').addEventListener('click', () => {
      switchSubView('step4', 'plans');
      renderImprovementPlans(getCurrentProject());
    });
    $('#btnStep4Gantt').addEventListener('click', () => {
      switchSubView('step4', 'gantt');
      renderGanttChart();
    });

    // Step5 サブビュー切り替え
    $('#btnStep5Measure').addEventListener('click', () => {
      switchSubView('step5', 'measure');
      renderMeasurements(getCurrentProject());
    });
    $('#btnStep5Report').addEventListener('click', () => {
      switchSubView('step5', 'report');
      renderReportPanel();
    });

    // 改善計画保存
    $('#btnSaveImprovement').addEventListener('click', saveImprovement);

    // 効果測定保存
    $('#btnSaveMeasure').addEventListener('click', saveMeasure);

    // 設定
    $('#btnAddCategory').addEventListener('click', () => {
      const val = $('#newCategoryInput').value.trim();
      if (!val) return;
      if (appData.processCategories.includes(val)) {
        showToast('同名のプロセス区分が既に存在します', 'error');
        return;
      }
      appData.processCategories.push(val);
      saveData(appData);
      $('#newCategoryInput').value = '';
      renderSettings();
      showToast('プロセス区分を追加しました', 'success');
    });

    $('#btnExportAll').addEventListener('click', exportAll);
    $('#btnImportAll').addEventListener('click', importAll);
    const btnExport = $('#btnExportProject');
    if (btnExport) btnExport.addEventListener('click', exportAll);

    // AIアシスト
    $('#btnAIGenerate').addEventListener('click', generateAISuggestion);
    $('#btnAIApply').addEventListener('click', applyAISuggestion);
  }

  function switchSubView(step, subName) {
    const viewEl = $(`#view-${step}`);
    if (!viewEl) return;

    viewEl.querySelectorAll('.sub-view').forEach(sv => {
      sv.classList.remove('active');
      sv.style.display = 'none';
    });
    const target = $(`#sub-${subName}`);
    if (target) {
      target.classList.add('active');
      target.style.display = '';
    }

    viewEl.querySelectorAll('.view-actions .btn').forEach(btn => {
      btn.classList.remove('active');
      if (btn.dataset.sub === subName) btn.classList.add('active');
    });
  }

  // ==========================================
  // ドロップダウンメニュー制御
  // ==========================================
  function setupDropdowns() {
    $$('.dropdown-toggle').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const dropdown = btn.closest('.dropdown');
        const menu = dropdown.querySelector('.dropdown-menu');
        const isOpen = menu.classList.contains('show');
        closeAllDropdowns();
        if (!isOpen) menu.classList.add('show');
      });
    });
    // 外側クリックで閉じる
    document.addEventListener('click', () => closeAllDropdowns());
  }

  function closeAllDropdowns() {
    $$('.dropdown-menu').forEach(m => m.classList.remove('show'));
  }

  /** 保存メニューの表示制御 */
  function updateSaveProjectVisibility() {
    const proj = getCurrentProject();
    const isStepView = currentView && currentView.startsWith('step');
    const btnCompany = $('#btnFileSaveCompany');
    const btnProj = $('#btnFileSaveProj');
    if (btnCompany) btnCompany.style.display = (proj && isStepView) ? '' : 'none';
    if (btnProj) btnProj.style.display = (proj && isStepView) ? '' : 'none';
  }

  // ==========================================
  // ファイル保存・読込
  // ==========================================

  /** JSONをダウンロードするヘルパー */
  function downloadJson(data, filename) {
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  /** ファイル読込ヘルパー */
  function readJsonFile(inputId, callback) {
    const input = $(inputId);
    input.click();
    input.onchange = () => {
      const file = input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = JSON.parse(e.target.result);
          callback(data, file.name);
        } catch (err) {
          showToast('ファイルの読み込みに失敗しました', 'error');
        }
      };
      reader.readAsText(file);
      input.value = '';
    };
  }

  function sanitizeFilename(str) {
    return (str || '').replace(/[()（）株\s\/\\:*?"<>|]/g, '');
  }

  const today = () => new Date().toISOString().split('T')[0];

  /** 全データ保存 */
  function fileSaveAll() {
    downloadJson(appData, `bpi_navi_全データ_${today()}.json`);
    showToast('全データをファイルに保存しました', 'success');
  }

  /** 企業単位で保存（同じ企業名の全プロジェクト） */
  function fileSaveCompany() {
    const proj = getCurrentProject();
    if (!proj) { showToast('プロジェクトが選択されていません', 'error'); return; }
    const company = proj.company;
    const companyProjects = appData.projects.filter(p => p.company === company);
    const companyInfo = (appData.companies || {})[company] || {};
    const exportData = {
      _type: 'bpi_company',
      _version: 1,
      company: company,
      companyInfo: companyInfo,
      projects: companyProjects
    };
    downloadJson(exportData, `bpi_company_${sanitizeFilename(company)}_${today()}.json`);
    showToast(`「${company}」の${companyProjects.length}件のプロジェクトを保存しました`, 'success');
  }

  /** プロジェクト単体を保存 */
  function fileSaveProject() {
    const proj = getCurrentProject();
    if (!proj) { showToast('プロジェクトが選択されていません', 'error'); return; }
    const exportData = {
      _type: 'bpi_project',
      _version: 1,
      project: proj
    };
    const fname = `bpi_project_${sanitizeFilename(proj.company)}_${sanitizeFilename(proj.name)}_${today()}.json`;
    downloadJson(exportData, fname);
    showToast(`「${proj.name}」を保存しました`, 'success');
  }

  /** 全データ読込（上書き） */
  function fileLoadAll() {
    readJsonFile('#loadFileInput', (imported, filename) => {
      if (imported.projects) {
        appData = migrateData(imported);
        saveData(appData);
        showToast(`${filename} を読み込みました（全データ上書き）`, 'success');
        location.reload();
      } else {
        showToast('無効なデータ形式です。全データ形式のJSONを選択してください', 'error');
      }
    });
  }

  /**
   * プロジェクト読込（更新/追加）
   * - 同じIDのプロジェクトが既にある → 上書き更新
   * - 同じIDがない → 新規追加
   * 企業ファイル・プロジェクト単体ファイル・全データファイルすべて対応
   */
  function fileLoadMerge() {
    readJsonFile('#loadProjectInput', (imported, filename) => {
      let projects = [];

      if (imported._type === 'bpi_company' && imported.projects) {
        // 企業単位ファイル
        projects = imported.projects;
        // 企業情報も取り込む
        if (imported.companyInfo && imported.company) {
          if (!appData.companies) appData.companies = {};
          appData.companies[imported.company] = imported.companyInfo;
        }
      } else if (imported._type === 'bpi_project' && imported.project) {
        // プロジェクト単体ファイル
        projects = [imported.project];
      } else if (imported.projects && imported.projects.length > 0) {
        // 全データ形式 — 企業情報もあれば取り込む
        projects = imported.projects;
        if (imported.companies) {
          if (!appData.companies) appData.companies = {};
          Object.assign(appData.companies, imported.companies);
        }
      }

      if (projects.length === 0) {
        showToast('無効なデータ形式です', 'error');
        return;
      }

      let updated = 0;
      let added = 0;

      projects.forEach(proj => {
        // プロジェクトの企業がcompaniesに未登録なら自動登録
        if (proj.company && !appData.companies[proj.company]) {
          appData.companies[proj.company] = { description: '' };
        }
        const idx = appData.projects.findIndex(p => p.id === proj.id);
        if (idx >= 0) {
          // 既存プロジェクトを上書き更新
          appData.projects[idx] = proj;
          updated++;
        } else {
          // 新規追加
          appData.projects.push(proj);
          added++;
        }
      });

      saveData(appData);

      const msgs = [];
      if (updated > 0) msgs.push(`${updated}件更新`);
      if (added > 0) msgs.push(`${added}件追加`);
      showToast(`${msgs.join('・')}しました`, 'success');
      navigate('dashboard');
    });
  }

  // ==========================================
  // デモデータ生成（(株)KK精工 - 工業用切削工具製造業）
  // プロセスごとに独立したプロジェクトとして生成
  // ==========================================
  function generateDemoData() {
    let n = 0;
    const tid = () => 'dt_' + (++n);

    // === 受注プロセス（16件） ===
    const ord = [
      { id: tid(), category: '受注', code: 'ORD-001', content: '顧客からの見積依頼受付（電話・FAX・メール）', person: '営業事務', target: '見積依頼書', method: 'FAX/電話/メールで受領し紙台帳に転記', timeRequired: '20', timeUnit: '分', freqType: '日次', freqCount: '5', tools: 'FAX, 電話, Excel', notes: 'FAXの場合が約6割。手書き記録が多い', scores: { timeImpact: 4, qualityImpact: 3, frequency: 5, difficulty: 2 }, problems: ['duplicate', 'paper'] },
      { id: tid(), category: '受注', code: 'ORD-002', content: '見積依頼内容の確認・工具仕様の精査', person: '営業担当', target: '依頼内容', method: '顧客の図面・仕様書と過去類似品を突合確認', timeRequired: '25', timeUnit: '分', freqType: '日次', freqCount: '4', tools: '図面, 製品カタログ', notes: '特殊工具は技術部門への確認が必要で時間がかかる', scores: { timeImpact: 3, qualityImpact: 4, frequency: 4, difficulty: 3 }, problems: ['waiting', 'personal'] },
      { id: tid(), category: '受注', code: 'ORD-003', content: '見積書作成（標準品の価格計算）', person: '営業担当', target: '見積書', method: '価格表Excelから単価を検索し見積書を作成', timeRequired: '30', timeUnit: '分', freqType: '日次', freqCount: '3', tools: 'Excel, 価格表', notes: '価格表が複数ファイルに分散し検索に時間がかかる', scores: { timeImpact: 4, qualityImpact: 3, frequency: 4, difficulty: 2 }, problems: ['search', 'nostandard'] },
      { id: tid(), category: '受注', code: 'ORD-004', content: '見積書作成（特殊形状工具の個別原価計算）', person: '営業担当', target: '見積書', method: '加工条件・材料費・工数を個別に積算しExcelで計算', timeRequired: '60', timeUnit: '分', freqType: '日次', freqCount: '2', tools: 'Excel, 原価計算シート', notes: 'ベテラン営業しか精度高く積算できない', scores: { timeImpact: 5, qualityImpact: 5, frequency: 4, difficulty: 4 }, problems: ['personal', 'search'] },
      { id: tid(), category: '受注', code: 'ORD-005', content: '見積書の上長承認取得', person: '営業担当', target: '見積書', method: '紙の見積書に上長の承認印をもらう', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '4', tools: '紙, 印鑑', notes: '上長不在時に承認待ちが発生（半日〜1日遅延）', scores: { timeImpact: 3, qualityImpact: 2, frequency: 4, difficulty: 1 }, problems: ['waiting', 'paper'] },
      { id: tid(), category: '受注', code: 'ORD-006', content: '見積書の顧客への送付', person: '営業事務', target: '見積書', method: 'FAXまたはメール添付（PDF化）で送付', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '4', tools: 'FAX, メール, PDF変換', notes: 'FAX送付先の確認に手間がかかる', scores: { timeImpact: 2, qualityImpact: 2, frequency: 4, difficulty: 1 }, problems: ['paper'] },
      { id: tid(), category: '受注', code: 'ORD-007', content: '注文書の受領・内容確認', person: '営業事務', target: '注文書', method: 'FAX/メール/郵送で受領し見積内容と照合', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '4', tools: 'FAX, 見積書控え', notes: '見積と注文の仕様差異が月3件程度発生', scores: { timeImpact: 3, qualityImpact: 4, frequency: 4, difficulty: 2 }, problems: ['duplicate', 'paper'] },
      { id: tid(), category: '受注', code: 'ORD-008', content: '受注情報の基幹システムへの入力', person: '営業事務', target: '受注データ', method: '紙の注文書を見て基幹システムに手入力', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '4', tools: '基幹システム', notes: '転記ミスによる納期・仕様間違いが月2件程度発生', scores: { timeImpact: 3, qualityImpact: 5, frequency: 4, difficulty: 2 }, problems: ['duplicate', 'paper'] },
      { id: tid(), category: '受注', code: 'ORD-009', content: '納期回答の確認（生産管理への問合せ）', person: '営業担当', target: '納期回答', method: '生産管理に口頭・電話で納期確認', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '4', tools: '電話, 内線', notes: '生産管理の回答を待つ時間が長い（平均2時間）', scores: { timeImpact: 4, qualityImpact: 3, frequency: 4, difficulty: 2 }, problems: ['waiting', 'communication'] },
      { id: tid(), category: '受注', code: 'ORD-010', content: '納期回答の顧客への連絡', person: '営業担当', target: '顧客', method: '電話またはメールで納期を連絡', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '4', tools: '電話, メール', notes: '短納期要請への対応が多い', scores: { timeImpact: 2, qualityImpact: 3, frequency: 4, difficulty: 1 }, problems: ['communication'] },
      { id: tid(), category: '受注', code: 'ORD-011', content: '受注残管理・進捗確認', person: '営業事務', target: '受注残リスト', method: 'Excelで受注残一覧を手動更新', timeRequired: '30', timeUnit: '分', freqType: '日次', freqCount: '1', tools: 'Excel', notes: '基幹システムとExcelの二重管理', scores: { timeImpact: 3, qualityImpact: 2, frequency: 3, difficulty: 2 }, problems: ['duplicate', 'system'] },
      { id: tid(), category: '受注', code: 'ORD-012', content: '仕様変更・設計変更の受付と展開', person: '営業担当', target: '変更依頼書', method: '顧客の変更要望を受け、社内関係部門に紙で展開', timeRequired: '20', timeUnit: '分', freqType: '週次', freqCount: '3', tools: '変更依頼書（紙）, メール', notes: '変更情報の伝達漏れによる誤加工が発生', scores: { timeImpact: 3, qualityImpact: 5, frequency: 3, difficulty: 3 }, problems: ['communication', 'paper'] },
      { id: tid(), category: '受注', code: 'ORD-013', content: '顧客別売上実績の集計・報告', person: '営業事務', target: '売上レポート', method: '基幹データとExcelで月次売上を集計', timeRequired: '60', timeUnit: '分', freqType: '月次', freqCount: '1', tools: 'Excel, 基幹システム', notes: 'データ抽出と加工に時間がかかる', scores: { timeImpact: 2, qualityImpact: 2, frequency: 2, difficulty: 2 }, problems: ['system', 'duplicate'] },
      { id: tid(), category: '受注', code: 'ORD-014', content: 'クレーム・返品受付と初動対応', person: '営業担当', target: 'クレーム報告', method: '顧客からの連絡を受け、品質管理に情報伝達', timeRequired: '30', timeUnit: '分', freqType: '週次', freqCount: '1', tools: '電話, メール, クレーム報告書', notes: '初動対応のスピードが顧客満足度に直結', scores: { timeImpact: 3, qualityImpact: 5, frequency: 2, difficulty: 3 }, problems: ['communication', 'nostandard'] },
      { id: tid(), category: '受注', code: 'ORD-015', content: '取引先マスタの更新・管理', person: '営業事務', target: '顧客マスタ', method: '新規取引先や住所変更を基幹システムに手入力', timeRequired: '15', timeUnit: '分', freqType: '週次', freqCount: '2', tools: '基幹システム', notes: 'マスタの不備により請求書送付先ミスが発生', scores: { timeImpact: 2, qualityImpact: 3, frequency: 2, difficulty: 1 }, problems: ['duplicate', 'nostandard'] },
      { id: tid(), category: '受注', code: 'ORD-016', content: 'リピート注文の自動引当・確認', person: '営業事務', target: '定期注文', method: '前回注文履歴を検索し、同仕様で受注処理', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '3', tools: '基幹システム, Excel', notes: 'リピート品の履歴検索に時間がかかる', scores: { timeImpact: 3, qualityImpact: 2, frequency: 4, difficulty: 2 }, problems: ['search'] }
    ];

    // === 生産管理プロセス（16件） ===
    const pln = [
      { id: tid(), category: '生産管理', code: 'PLN-001', content: '受注情報の確認と生産要否判断', person: '生産管理主任', target: '受注データ', method: '基幹システムの受注残を確認し、在庫引当か新規生産か判断', timeRequired: '20', timeUnit: '分', freqType: '日次', freqCount: '2', tools: '基幹システム, Excel', notes: '在庫状況と受注情報の突合が手作業', scores: { timeImpact: 3, qualityImpact: 3, frequency: 4, difficulty: 3 }, problems: ['system', 'personal'] },
      { id: tid(), category: '生産管理', code: 'PLN-002', content: '月次生産計画の立案', person: '生産管理主任', target: '月間生産計画表', method: '受注残・見込み・在庫をExcelで集約し計画立案', timeRequired: '180', timeUnit: '分', freqType: '月次', freqCount: '1', tools: 'Excel, ホワイトボード', notes: '担当者の経験と勘に依存。不在時に計画が立てられない', scores: { timeImpact: 5, qualityImpact: 4, frequency: 2, difficulty: 5 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '生産管理', code: 'PLN-003', content: '週次生産スケジュールの作成', person: '生産管理主任', target: '週間スケジュール', method: '月次計画をブレイクダウンし機械別にスケジュール作成', timeRequired: '120', timeUnit: '分', freqType: '週次', freqCount: '1', tools: 'Excel, ホワイトボード', notes: '機械の故障や急ぎ注文で頻繁に計画変更が発生', scores: { timeImpact: 5, qualityImpact: 4, frequency: 3, difficulty: 5 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '生産管理', code: 'PLN-004', content: '製造指示書の作成', person: '生産管理', target: '製造指示書', method: 'Excelテンプレートに品番・数量・納期・加工条件を入力', timeRequired: '20', timeUnit: '分', freqType: '日次', freqCount: '3', tools: 'Excel, プリンター', notes: '図面との紐づけが手作業で紛失リスクあり', scores: { timeImpact: 3, qualityImpact: 4, frequency: 4, difficulty: 2 }, problems: ['paper', 'nostandard'] },
      { id: tid(), category: '生産管理', code: 'PLN-005', content: '製造指示書・図面の現場配布', person: '生産管理', target: '製造指示書, 図面', method: '印刷した指示書と図面をセットにして各工程に配布', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '2', tools: 'プリンター', notes: '最新版管理が課題。旧版図面での加工ミスが年2回発生', scores: { timeImpact: 2, qualityImpact: 4, frequency: 4, difficulty: 2 }, problems: ['paper', 'nostandard'] },
      { id: tid(), category: '生産管理', code: 'PLN-006', content: '加工進捗の確認・実績収集', person: '生産管理', target: '進捗データ', method: '現場巡回で口頭確認、紙日報を回収しExcelに記録', timeRequired: '40', timeUnit: '分', freqType: '日次', freqCount: '2', tools: '日報用紙, Excel', notes: '紙の日報を集約するのに時間がかかる', scores: { timeImpact: 4, qualityImpact: 3, frequency: 5, difficulty: 3 }, problems: ['paper', 'waiting'] },
      { id: tid(), category: '生産管理', code: 'PLN-007', content: '機械稼働状況の把握', person: '生産管理', target: '機械稼働データ', method: '現場の稼働状況を目視確認し稼働率を計算', timeRequired: '20', timeUnit: '分', freqType: '日次', freqCount: '2', tools: '目視, Excel', notes: '機械の実稼働時間が正確に把握できていない', scores: { timeImpact: 3, qualityImpact: 2, frequency: 4, difficulty: 3 }, problems: ['paper', 'nostandard'] },
      { id: tid(), category: '生産管理', code: 'PLN-008', content: '納期遅延リスクの早期検知・対策', person: '生産管理主任', target: '納期管理表', method: '受注残と進捗を照合し遅延リスクを洗い出し', timeRequired: '30', timeUnit: '分', freqType: '日次', freqCount: '1', tools: 'Excel', notes: '遅延が発覚してから対策するケースが多い', scores: { timeImpact: 4, qualityImpact: 4, frequency: 4, difficulty: 3 }, problems: ['system', 'personal'] },
      { id: tid(), category: '生産管理', code: 'PLN-009', content: '特急品・割込み注文の計画調整', person: '生産管理主任', target: '生産計画', method: '既存計画の優先順位を手動で入れ替え調整', timeRequired: '30', timeUnit: '分', freqType: '日次', freqCount: '1', tools: 'Excel, ホワイトボード', notes: '特急品が入ると全体計画の見直しが必要', scores: { timeImpact: 4, qualityImpact: 3, frequency: 4, difficulty: 4 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '生産管理', code: 'PLN-010', content: '材料・超硬素材の在庫確認', person: '生産管理', target: '在庫リスト', method: '倉庫を目視確認し在庫台帳に記入', timeRequired: '25', timeUnit: '分', freqType: '日次', freqCount: '1', tools: '在庫台帳（紙）, Excel', notes: '在庫の数量差異が月に数回発生', scores: { timeImpact: 3, qualityImpact: 3, frequency: 4, difficulty: 3 }, problems: ['paper', 'search'] },
      { id: tid(), category: '生産管理', code: 'PLN-011', content: '材料の発注依頼（購買への連絡）', person: '生産管理', target: '発注依頼書', method: '不足材料を確認し購買担当にメモまたは口頭で依頼', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '1', tools: '紙メモ, 口頭', notes: '発注漏れによる材料欠品が月1回程度発生', scores: { timeImpact: 3, qualityImpact: 4, frequency: 4, difficulty: 2 }, problems: ['communication', 'nostandard'] },
      { id: tid(), category: '生産管理', code: 'PLN-012', content: 'コーティング外注の手配・納期管理', person: '生産管理', target: '外注依頼品', method: 'コーティング業者にFAXで依頼、納期確認', timeRequired: '20', timeUnit: '分', freqType: '週次', freqCount: '3', tools: 'FAX, 電話', notes: 'FAX送信と確認の手間。外注リードタイムのばらつき', scores: { timeImpact: 2, qualityImpact: 2, frequency: 3, difficulty: 2 }, problems: ['paper', 'waiting'] },
      { id: tid(), category: '生産管理', code: 'PLN-013', content: '完成品の入出庫管理', person: '生産管理', target: '在庫データ', method: '完成品を棚に格納し、出荷時に払い出し記録', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '2', tools: '在庫台帳（紙）, 棚', notes: '製品の所在が不明になることがある', scores: { timeImpact: 2, qualityImpact: 3, frequency: 4, difficulty: 2 }, problems: ['paper', 'search'] },
      { id: tid(), category: '生産管理', code: 'PLN-014', content: '生産実績の集計・月次報告', person: '生産管理', target: '月次生産報告書', method: '日次データをExcelで集計し月次レポートを作成', timeRequired: '90', timeUnit: '分', freqType: '月次', freqCount: '1', tools: 'Excel', notes: '日報データの集約に時間がかかる', scores: { timeImpact: 3, qualityImpact: 2, frequency: 2, difficulty: 2 }, problems: ['paper', 'duplicate'] },
      { id: tid(), category: '生産管理', code: 'PLN-015', content: '設備保全計画の管理', person: '生産管理', target: '保全計画表', method: 'Excelで定期保全の予定を管理', timeRequired: '30', timeUnit: '分', freqType: '月次', freqCount: '1', tools: 'Excel', notes: '保全漏れによる突発故障が年数回発生', scores: { timeImpact: 2, qualityImpact: 4, frequency: 2, difficulty: 3 }, problems: ['system', 'nostandard'] },
      { id: tid(), category: '生産管理', code: 'PLN-016', content: '工程間の仕掛品管理・滞留チェック', person: '生産管理', target: '仕掛品', method: '現場を巡回し仕掛品の滞留状況を確認', timeRequired: '20', timeUnit: '分', freqType: '日次', freqCount: '1', tools: '目視', notes: '仕掛品の長期滞留に気づかないことがある', scores: { timeImpact: 2, qualityImpact: 3, frequency: 4, difficulty: 2 }, problems: ['search', 'nostandard'] }
    ];

    // === 製造プロセス（17件） ===
    const mfg = [
      { id: tid(), category: '製造', code: 'MFG-001', content: 'CNC工具研削盤の段取り・プログラム呼出', person: 'CNCオペレータ', target: 'CNC加工プログラム', method: '図面を見てプログラムを検索・選択し治具をセット', timeRequired: '30', timeUnit: '分', freqType: '日次', freqCount: '4', tools: 'CNC研削盤, CAM', notes: '類似品プログラムの検索に時間がかかる', scores: { timeImpact: 4, qualityImpact: 4, frequency: 5, difficulty: 4 }, problems: ['search', 'personal'] },
      { id: tid(), category: '製造', code: 'MFG-002', content: 'CNCプログラムの新規作成・修正', person: 'プログラマ', target: 'NCプログラム', method: 'CAMソフトで工具形状に合わせたプログラムを作成', timeRequired: '60', timeUnit: '分', freqType: '日次', freqCount: '2', tools: 'CAMソフト', notes: '特殊形状はプログラム作成に半日以上かかることも', scores: { timeImpact: 4, qualityImpact: 4, frequency: 3, difficulty: 5 }, problems: ['personal'] },
      { id: tid(), category: '製造', code: 'MFG-003', content: '初品加工・試し削り', person: 'CNCオペレータ', target: '初品', method: '新規品や設定変更後に試し削りを行い寸法確認', timeRequired: '20', timeUnit: '分', freqType: '日次', freqCount: '3', tools: 'CNC研削盤, 測定器', notes: '初品OKまでの調整回数が品質に直結', scores: { timeImpact: 3, qualityImpact: 5, frequency: 4, difficulty: 3 }, problems: ['nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-004', content: '切削工具の研削加工（本加工）', person: 'CNCオペレータ', target: '加工済み工具', method: 'CNC工具研削盤による自動研削', timeRequired: '60', timeUnit: '分', freqType: '日次', freqCount: '6', tools: 'CNC工具研削盤', notes: '機械稼働中の異音・振動監視が必要', scores: { timeImpact: 2, qualityImpact: 2, frequency: 5, difficulty: 3 }, problems: [] },
      { id: tid(), category: '製造', code: 'MFG-005', content: '砥石のドレッシング・交換', person: 'CNCオペレータ', target: '砥石', method: '砥石の摩耗状態を確認しドレッシングまたは交換', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '2', tools: 'ドレッシングツール', notes: '交換タイミングの判断が経験依存', scores: { timeImpact: 2, qualityImpact: 4, frequency: 4, difficulty: 3 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-006', content: '再研削品の受入・摩耗状態確認', person: '製造リーダー', target: '使用済み工具', method: '返却工具の摩耗状態を目視・マイクロスコープで確認', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '3', tools: 'マイクロスコープ, 検査シート', notes: '再研削可否の判断基準が人によって異なる', scores: { timeImpact: 2, qualityImpact: 4, frequency: 4, difficulty: 4 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-007', content: '手仕上げ加工（バリ取り・研磨）', person: '仕上げ担当', target: '加工済み工具', method: '手作業でバリ取り・刃先処理・鏡面研磨', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '8', tools: '砥石, 研磨材', notes: '仕上がりの品質にばらつきが出やすい', scores: { timeImpact: 3, qualityImpact: 4, frequency: 5, difficulty: 3 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-008', content: '工程内寸法チェック（中間検査）', person: 'CNCオペレータ', target: '加工中の工具', method: '加工途中で寸法を測定し公差内か確認', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '10', tools: 'マイクロメータ, ノギス', notes: '全数チェックのため工数が大きい', scores: { timeImpact: 3, qualityImpact: 4, frequency: 5, difficulty: 2 }, problems: ['paper'] },
      { id: tid(), category: '製造', code: 'MFG-009', content: '刻印・マーキング作業', person: '仕上げ担当', target: '完成品', method: '製品に品番・ロットNo.をレーザーまたは手打ちで刻印', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '6', tools: 'レーザー刻印機, ポンチ', notes: '刻印ミス（品番間違い）が月1件程度', scores: { timeImpact: 1, qualityImpact: 3, frequency: 4, difficulty: 1 }, problems: ['nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-010', content: '治具・工具のセットアップ・保管管理', person: 'CNCオペレータ', target: '治具・工具', method: '使用する治具の選定・セット、使用後の清掃・返却', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '4', tools: '治具棚, 工具箱', notes: '治具の所在が不明で探す時間が発生', scores: { timeImpact: 3, qualityImpact: 2, frequency: 5, difficulty: 2 }, problems: ['search'] },
      { id: tid(), category: '製造', code: 'MFG-011', content: '作業日報の記入・提出', person: '各作業者', target: '作業日報', method: '紙の日報に手書きで作業実績・数量・不良数を記録', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '8', tools: '日報用紙', notes: '記入漏れ・後日まとめ書きが多い', scores: { timeImpact: 3, qualityImpact: 3, frequency: 5, difficulty: 1 }, problems: ['paper', 'nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-012', content: '設備の日常点検・始業点検', person: 'CNCオペレータ', target: 'CNC研削盤', method: '始業前に油量・エア圧・各部動作を確認し点検表記入', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '4', tools: '点検表（紙）', notes: '点検の形骸化（チェックだけで実際の確認不十分）', scores: { timeImpact: 1, qualityImpact: 3, frequency: 5, difficulty: 1 }, problems: ['paper', 'nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-013', content: '不良品発生時の原因調査・対処', person: '製造リーダー', target: '不良品', method: '不良内容を確認し原因を特定、再加工または廃棄判断', timeRequired: '30', timeUnit: '分', freqType: '週次', freqCount: '3', tools: '測定器, 不良報告書', notes: '原因特定に時間がかかる。同じ不良の再発あり', scores: { timeImpact: 3, qualityImpact: 5, frequency: 3, difficulty: 4 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-014', content: '5S活動・職場清掃', person: '各作業者', target: '作業場', method: '作業エリアの清掃・整理整頓', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '1', tools: '清掃用具', notes: '切粉の清掃が重要。清掃不足は品質に影響', scores: { timeImpact: 1, qualityImpact: 2, frequency: 4, difficulty: 1 }, problems: [] },
      { id: tid(), category: '製造', code: 'MFG-015', content: '作業標準書の確認・更新', person: '製造リーダー', target: '作業標準書', method: '新製品や工程変更時に作業標準書を改訂', timeRequired: '45', timeUnit: '分', freqType: '月次', freqCount: '2', tools: 'Word, 紙ファイル', notes: '標準書が最新化されていないケースがある', scores: { timeImpact: 2, qualityImpact: 4, frequency: 2, difficulty: 3 }, problems: ['paper', 'personal'] },
      { id: tid(), category: '製造', code: 'MFG-016', content: '新人・多能工の技能教育（OJT）', person: '製造リーダー', target: '作業者', method: 'マンツーマンで加工技術を指導', timeRequired: '60', timeUnit: '分', freqType: '週次', freqCount: '3', tools: '実機, 作業標準書', notes: '教育記録が残っていない。技能マップが未整備', scores: { timeImpact: 3, qualityImpact: 4, frequency: 3, difficulty: 4 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '製造', code: 'MFG-017', content: '加工条件データの記録・蓄積', person: 'CNCオペレータ', target: '加工条件表', method: '上手くいった加工条件をノートに手書きメモ', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '3', tools: 'ノート', notes: '個人ノートのため共有されない。退職時にノウハウ喪失', scores: { timeImpact: 2, qualityImpact: 4, frequency: 4, difficulty: 2 }, problems: ['personal', 'paper'] }
    ];

    // === 品質管理プロセス（16件） ===
    const qc = [
      { id: tid(), category: '品質管理', code: 'QC-001', content: '受入検査（材料・素材の検査）', person: '検査員', target: '入荷材料', method: '納品された超硬素材の外観・寸法・ミルシート確認', timeRequired: '15', timeUnit: '分', freqType: '週次', freqCount: '4', tools: 'ノギス, ミルシート', notes: 'ミルシートの保管が煩雑', scores: { timeImpact: 2, qualityImpact: 4, frequency: 3, difficulty: 2 }, problems: ['paper'] },
      { id: tid(), category: '品質管理', code: 'QC-002', content: '工程内検査（寸法測定）', person: '検査員', target: '加工済み工具', method: '測定器具で外径・全長・刃先角度等を測定し記録', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '10', tools: 'マイクロメータ, 投影機, 検査記録', notes: '測定データの手書き記録→Excelへの転記', scores: { timeImpact: 4, qualityImpact: 5, frequency: 5, difficulty: 3 }, problems: ['duplicate', 'paper'] },
      { id: tid(), category: '品質管理', code: 'QC-003', content: '最終検査（出荷前検査）', person: '検査員', target: '出荷予定品', method: '全数または抜取りで最終寸法・外観検査を実施', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '8', tools: '測定器具一式', notes: '検査項目が多い製品は時間がかかる', scores: { timeImpact: 3, qualityImpact: 5, frequency: 5, difficulty: 2 }, problems: ['nostandard'] },
      { id: tid(), category: '品質管理', code: 'QC-004', content: '検査成績書の作成', person: '検査員', target: '検査成績書', method: '測定データをExcelの顧客指定フォーマットに転記', timeRequired: '20', timeUnit: '分', freqType: '日次', freqCount: '5', tools: 'Excel, プリンター', notes: '顧客ごとにフォーマットが異なる（約15種類）', scores: { timeImpact: 3, qualityImpact: 4, frequency: 4, difficulty: 2 }, problems: ['duplicate', 'nostandard'] },
      { id: tid(), category: '品質管理', code: 'QC-005', content: '出荷判定の承認', person: '品質管理責任者', target: '検査成績書', method: '検査結果を確認し出荷可否を判定・押印', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '5', tools: '検査成績書, 印鑑', notes: '責任者不在時の代行ルールが不明確', scores: { timeImpact: 2, qualityImpact: 3, frequency: 4, difficulty: 1 }, problems: ['waiting', 'nostandard'] },
      { id: tid(), category: '品質管理', code: 'QC-006', content: '不適合品の識別・隔離', person: '検査員', target: '不適合品', method: '不合格品に赤タグを貼り隔離棚に移動', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '2', tools: '赤タグ, 隔離棚', notes: '隔離が不十分で不適合品が流出するリスク', scores: { timeImpact: 1, qualityImpact: 5, frequency: 4, difficulty: 1 }, problems: ['nostandard'] },
      { id: tid(), category: '品質管理', code: 'QC-007', content: '不適合報告書の作成・是正処置', person: '品質管理責任者', target: '不適合報告書', method: '不適合内容・原因・是正措置を紙の報告書に記録', timeRequired: '45', timeUnit: '分', freqType: '週次', freqCount: '2', tools: '不適合報告書（紙）', notes: 'ISO9001要求の記録。過去事例の検索が困難', scores: { timeImpact: 3, qualityImpact: 5, frequency: 3, difficulty: 3 }, problems: ['paper', 'search'] },
      { id: tid(), category: '品質管理', code: 'QC-008', content: '顧客クレームの調査・回答書作成', person: '品質管理責任者', target: 'クレーム回答書', method: '不良原因を調査し再発防止策を含めた回答書を作成', timeRequired: '120', timeUnit: '分', freqType: '月次', freqCount: '2', tools: 'Word, 測定器具', notes: '調査に時間がかかり回答期限に追われる', scores: { timeImpact: 4, qualityImpact: 5, frequency: 2, difficulty: 4 }, problems: ['personal', 'search'] },
      { id: tid(), category: '品質管理', code: 'QC-009', content: '品質データの集計・分析（月次）', person: '品質管理', target: '品質月報', method: '不良率・クレーム件数等をExcelで集計しグラフ化', timeRequired: '60', timeUnit: '分', freqType: '月次', freqCount: '1', tools: 'Excel', notes: 'データ入力から集計まで全て手作業', scores: { timeImpact: 3, qualityImpact: 3, frequency: 2, difficulty: 2 }, problems: ['paper', 'duplicate'] },
      { id: tid(), category: '品質管理', code: 'QC-010', content: '測定機器の校正管理', person: '品質管理', target: '校正台帳', method: 'Excelの校正台帳で期限管理、紙の校正証明書保管', timeRequired: '30', timeUnit: '分', freqType: '月次', freqCount: '1', tools: 'Excel, 校正証明書', notes: '期限超過に気づかないリスクあり', scores: { timeImpact: 1, qualityImpact: 4, frequency: 1, difficulty: 2 }, problems: ['paper', 'system'] },
      { id: tid(), category: '品質管理', code: 'QC-011', content: '測定機器の日常点検', person: '検査員', target: '測定器具', method: '使用前にゲージブロックで測定器の精度確認', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '2', tools: 'ゲージブロック, 点検表', notes: '点検記録が形骸化している', scores: { timeImpact: 1, qualityImpact: 3, frequency: 4, difficulty: 1 }, problems: ['paper', 'nostandard'] },
      { id: tid(), category: '品質管理', code: 'QC-012', content: 'ISO9001内部監査の実施', person: '品質管理責任者', target: '各部門', method: '年間計画に基づき各部門の監査を実施・報告', timeRequired: '480', timeUnit: '分', freqType: '月次', freqCount: '1', tools: '監査チェックリスト', notes: '監査準備と報告書作成に大きな工数', scores: { timeImpact: 3, qualityImpact: 3, frequency: 1, difficulty: 4 }, problems: ['paper', 'personal'] },
      { id: tid(), category: '品質管理', code: 'QC-013', content: 'QC工程表・管理計画書の維持管理', person: '品質管理', target: 'QC工程表', method: '新製品や工程変更時にQC工程表を作成・改訂', timeRequired: '60', timeUnit: '分', freqType: '月次', freqCount: '2', tools: 'Excel, 紙', notes: '最新版管理が不十分', scores: { timeImpact: 2, qualityImpact: 4, frequency: 2, difficulty: 3 }, problems: ['paper', 'nostandard'] },
      { id: tid(), category: '品質管理', code: 'QC-014', content: '工程能力の評価（Cp/Cpk計算）', person: '品質管理', target: '測定データ', method: '測定データからExcelで工程能力指数を計算', timeRequired: '30', timeUnit: '分', freqType: '月次', freqCount: '2', tools: 'Excel', notes: 'データの蓄積と分析が手作業で非効率', scores: { timeImpact: 2, qualityImpact: 3, frequency: 2, difficulty: 3 }, problems: ['system', 'personal'] },
      { id: tid(), category: '品質管理', code: 'QC-015', content: '品質記録の文書管理・保管', person: '品質管理', target: '品質記録一式', method: '検査記録・不適合報告等を紙ファイルで保管', timeRequired: '20', timeUnit: '分', freqType: '日次', freqCount: '1', tools: '紙ファイル, キャビネット', notes: '保管スペースの不足。過去記録の検索に時間がかかる', scores: { timeImpact: 2, qualityImpact: 2, frequency: 4, difficulty: 1 }, problems: ['paper', 'search'] },
      { id: tid(), category: '品質管理', code: 'QC-016', content: '仕入先の品質評価・監査', person: '品質管理責任者', target: '仕入先', method: '仕入先の品質実績を集計し評価シートを作成', timeRequired: '60', timeUnit: '分', freqType: '月次', freqCount: '1', tools: 'Excel, 評価シート', notes: '評価基準が明文化されていない', scores: { timeImpact: 2, qualityImpact: 3, frequency: 1, difficulty: 3 }, problems: ['personal', 'nostandard'] }
    ];

    // === 出荷プロセス（15件） ===
    const shp = [
      { id: tid(), category: '出荷', code: 'SHP-001', content: '出荷指示書の受領・内容確認', person: '出荷担当', target: '出荷指示書', method: '生産管理からの出荷指示書を受領し、品番・数量・納期を確認', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '5', tools: '出荷指示書（紙）', notes: '指示書の記載ミスが月2回程度発生', scores: { timeImpact: 3, qualityImpact: 4, frequency: 5, difficulty: 2 }, problems: ['paper', 'duplicate'] },
      { id: tid(), category: '出荷', code: 'SHP-002', content: '出荷対象品の在庫引当・ピッキング', person: '出荷担当', target: '完成品', method: '完成品棚から該当品をピッキングし出荷エリアに集約', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '5', tools: '在庫台帳, 棚番表', notes: '棚番が分かりにくく探す時間が発生。品番間違いも月1件', scores: { timeImpact: 3, qualityImpact: 4, frequency: 5, difficulty: 2 }, problems: ['search', 'nostandard'] },
      { id: tid(), category: '出荷', code: 'SHP-003', content: '出荷前の最終数量確認（検品）', person: '出荷担当', target: '出荷品', method: '出荷指示書と現物を照合し数量・品番を確認', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '5', tools: '出荷指示書, ノギス', notes: '数量違いの出荷が月1件程度発生', scores: { timeImpact: 2, qualityImpact: 5, frequency: 5, difficulty: 1 }, problems: ['nostandard'] },
      { id: tid(), category: '出荷', code: 'SHP-004', content: '製品の梱包作業', person: '出荷担当', target: '出荷品', method: '製品サイズに合わせた箱選定・緩衝材配置・梱包', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '5', tools: '段ボール, 緩衝材, テープ', notes: '切削工具は刃先保護が重要。破損クレームが年3件', scores: { timeImpact: 3, qualityImpact: 4, frequency: 5, difficulty: 2 }, problems: ['nostandard', 'personal'] },
      { id: tid(), category: '出荷', code: 'SHP-005', content: '送り状・納品書の作成', person: '出荷事務', target: '送り状, 納品書', method: '基幹システムから出荷データを出力し、送り状を手入力で作成', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '5', tools: '基幹システム, 送り状発行ソフト', notes: '基幹と送り状ソフトが連携していないため二重入力', scores: { timeImpact: 4, qualityImpact: 3, frequency: 5, difficulty: 2 }, problems: ['duplicate', 'system'] },
      { id: tid(), category: '出荷', code: 'SHP-006', content: '運送業者の手配・集荷依頼', person: '出荷事務', target: '運送業者', method: '電話またはWebで運送業者に集荷を依頼', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '3', tools: '電話, 運送業者Web', notes: '急ぎの場合の手配に時間がかかる', scores: { timeImpact: 2, qualityImpact: 2, frequency: 4, difficulty: 1 }, problems: ['communication'] },
      { id: tid(), category: '出荷', code: 'SHP-007', content: '出荷実績の基幹システム入力', person: '出荷事務', target: '出荷データ', method: '出荷完了後に基幹システムに出荷実績を手入力', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '5', tools: '基幹システム', notes: '入力忘れにより在庫数に差異が発生することがある', scores: { timeImpact: 3, qualityImpact: 4, frequency: 5, difficulty: 2 }, problems: ['duplicate', 'paper'] },
      { id: tid(), category: '出荷', code: 'SHP-008', content: '出荷案内の顧客への連絡', person: '出荷事務', target: '顧客', method: 'メールまたはFAXで出荷案内（送り状番号）を送付', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '5', tools: 'メール, FAX', notes: '連絡漏れにより顧客からの問合せが発生', scores: { timeImpact: 2, qualityImpact: 3, frequency: 5, difficulty: 1 }, problems: ['communication', 'paper'] },
      { id: tid(), category: '出荷', code: 'SHP-009', content: '納品書控え・出荷記録の保管', person: '出荷事務', target: '納品書控え', method: '紙の納品書控えをファイリングして保管', timeRequired: '5', timeUnit: '分', freqType: '日次', freqCount: '5', tools: '紙ファイル', notes: '保管スペース不足。過去の出荷記録検索に時間がかかる', scores: { timeImpact: 2, qualityImpact: 2, frequency: 5, difficulty: 1 }, problems: ['paper', 'search'] },
      { id: tid(), category: '出荷', code: 'SHP-010', content: '配送状況の追跡・顧客問合せ対応', person: '出荷事務', target: '配送状況', method: '運送業者のWebサイトで追跡番号を照会し顧客に回答', timeRequired: '10', timeUnit: '分', freqType: '日次', freqCount: '3', tools: '運送業者Web, 電話', notes: '追跡番号の管理がExcelで煩雑', scores: { timeImpact: 2, qualityImpact: 3, frequency: 4, difficulty: 1 }, problems: ['search', 'system'] },
      { id: tid(), category: '出荷', code: 'SHP-011', content: '返品・交換品の受入処理', person: '出荷担当', target: '返品', method: '返品の受入検査・在庫戻し処理・報告書作成', timeRequired: '30', timeUnit: '分', freqType: '週次', freqCount: '2', tools: '検査器具, 返品報告書', notes: '返品理由の分析ができていない', scores: { timeImpact: 2, qualityImpact: 4, frequency: 2, difficulty: 3 }, problems: ['nostandard', 'paper'] },
      { id: tid(), category: '出荷', code: 'SHP-012', content: '月次出荷実績の集計・報告', person: '出荷事務', target: '出荷レポート', method: '基幹データとExcelで月次出荷数量・金額を集計', timeRequired: '60', timeUnit: '分', freqType: '月次', freqCount: '1', tools: 'Excel, 基幹システム', notes: 'データ抽出と集計が手作業で時間がかかる', scores: { timeImpact: 3, qualityImpact: 2, frequency: 2, difficulty: 2 }, problems: ['duplicate', 'system'] },
      { id: tid(), category: '出荷', code: 'SHP-013', content: '梱包資材の在庫管理・発注', person: '出荷担当', target: '梱包資材', method: '段ボール・緩衝材等の在庫を目視確認し不足時に発注', timeRequired: '15', timeUnit: '分', freqType: '週次', freqCount: '1', tools: '目視, 発注書', notes: '資材切れで出荷が遅れることがある', scores: { timeImpact: 2, qualityImpact: 2, frequency: 3, difficulty: 1 }, problems: ['search', 'nostandard'] },
      { id: tid(), category: '出荷', code: 'SHP-014', content: '輸出品の通関書類作成', person: '出荷事務', target: 'インボイス, パッキングリスト', method: '輸出用の書類（インボイス・パッキングリスト）を手作成', timeRequired: '45', timeUnit: '分', freqType: '週次', freqCount: '2', tools: 'Excel, 通関書類テンプレート', notes: 'HSコードの確認に時間がかかる。記載ミスで通関遅延リスク', scores: { timeImpact: 4, qualityImpact: 5, frequency: 3, difficulty: 4 }, problems: ['personal', 'nostandard'] },
      { id: tid(), category: '出荷', code: 'SHP-015', content: '出荷スケジュールの調整・優先度管理', person: '出荷担当', target: '出荷スケジュール', method: '営業・生産管理と調整し出荷優先順位を決定', timeRequired: '15', timeUnit: '分', freqType: '日次', freqCount: '1', tools: '電話, ホワイトボード', notes: '特急出荷の割込みで通常出荷が遅延することがある', scores: { timeImpact: 3, qualityImpact: 3, frequency: 4, difficulty: 3 }, problems: ['communication', 'waiting'] }
    ];

    // --- ヘルパー関数 ---
    const allTasks = [...ord, ...pln, ...mfg, ...qc, ...shp];
    allTasks.forEach(t => { t.ecrs = null; });

    // ムリ・ムダ・ムラをproblemsから自動分類
    const mmmMap = {
      waiting: ['ムリ', 'ムダ'], paper: ['ムダ'], duplicate: ['ムダ'],
      communication: ['ムラ'], search: ['ムダ'], system: ['ムラ'],
      nostandard: ['ムラ'], personal: ['ムリ']
    };
    allTasks.forEach(t => {
      const set = new Set();
      (t.problems || []).forEach(p => (mmmMap[p] || []).forEach(m => set.add(m)));
      t.mmm = [...set];
    });
    const byCode = (code) => { const t = allTasks.find(x => x.code === code); return t ? t.id : null; };

    // --- ECRS判定をタスクに適用 ---
    const applyEcrs = (list) => {
      list.forEach(e => {
        const t = allTasks.find(x => x.code === e.code);
        if (t) t.ecrs = e.ecrs;
      });
    };

    // 受注ECRS（Step 5: 全完了）
    applyEcrs([
      { code: 'ORD-001', ecrs: { E: 'no', C: 'maybe', R: 'no', S: 'yes' } },
      { code: 'ORD-004', ecrs: { E: 'no', C: 'no', R: 'no', S: 'yes' } },
      { code: 'ORD-008', ecrs: { E: 'no', C: 'yes', R: 'no', S: 'yes' } },
      { code: 'ORD-009', ecrs: { E: 'no', C: 'no', R: 'yes', S: 'maybe' } },
      { code: 'ORD-011', ecrs: { E: 'yes', C: 'no', R: 'no', S: 'no' } }
    ]);

    // 生産管理ECRS（Step 4: 改善策検討中）
    applyEcrs([
      { code: 'PLN-002', ecrs: { E: 'no', C: 'no', R: 'no', S: 'yes' } },
      { code: 'PLN-003', ecrs: { E: 'no', C: 'no', R: 'no', S: 'yes' } },
      { code: 'PLN-006', ecrs: { E: 'no', C: 'yes', R: 'no', S: 'yes' } },
      { code: 'PLN-008', ecrs: { E: 'no', C: 'no', R: 'no', S: 'yes' } },
      { code: 'PLN-009', ecrs: { E: 'no', C: 'no', R: 'yes', S: 'maybe' } }
    ]);

    // 製造ECRS（Step 3: ECRS分析中）
    applyEcrs([
      { code: 'MFG-001', ecrs: { E: 'no', C: 'no', R: 'no', S: 'yes' } },
      { code: 'MFG-006', ecrs: { E: 'no', C: 'no', R: 'no', S: 'yes' } },
      { code: 'MFG-011', ecrs: { E: 'no', C: 'yes', R: 'no', S: 'yes' } },
      { code: 'MFG-017', ecrs: { E: 'no', C: 'no', R: 'no', S: 'yes' } }
    ]);
    // 品質管理・出荷はECRS未実施

    // --- 受注プロジェクト: 改善計画4件 + 測定2件 ---
    const ordImprovements = [
      { taskId: byCode('ORD-001'), ecrsKey: 'S', content: 'Web見積依頼フォームの導入', method: 'Googleフォーム＋スプレッドシートで見積依頼を受付', person: '営業部長', resources: 'Googleフォーム（無料）、連携設定5万円', startDate: '2026-01-15', endDate: '2026-03-31', effectQuantity: '見積受付工数 月40時間→15時間（63%削減）', effectQuality: '転記ミスゼロ化', status: 'done' },
      { taskId: byCode('ORD-004'), ecrsKey: 'S', content: '見積自動計算シートの整備', method: '過去見積実績DBを構築し類似品検索・価格自動算出', person: '営業担当', resources: 'Excel VBA改修（社内対応）', startDate: '2026-02-01', endDate: '2026-04-30', effectQuantity: '見積作成時間 60分→20分/件', effectQuality: '見積精度の向上・属人化解消', status: 'done' },
      { taskId: byCode('ORD-008'), ecrsKey: 'C', content: '見積→受注の自動データ連携', method: '見積承認時に受注データを基幹に自動転送', person: 'システム担当', resources: 'RPA導入費15万円', startDate: '2026-03-01', endDate: '2026-05-31', effectQuantity: '受注入力工数 月20時間→5時間', effectQuality: '転記ミスによる納期間違いゼロ化', status: 'done' },
      { taskId: byCode('ORD-011'), ecrsKey: 'E', content: 'Excel受注残台帳の廃止（基幹一本化）', method: '基幹の受注残管理機能を活用しExcel二重管理を廃止', person: '営業事務', resources: '追加費用なし', startDate: '2026-02-01', endDate: '2026-03-31', effectQuantity: '台帳管理 月10時間→0時間', effectQuality: '数値不一致の解消', status: 'done' }
    ];
    const ordMeasurements = [
      { taskId: byCode('ORD-001'), ecrsKey: 'S', afterTime: '8', achievement: 'Web依頼フォーム導入済。FAX比率60%→15%に低下。転記ミスは対象期間ゼロ件', date: '2026-03-05', comment: '顧客の一部は未だFAX希望。段階的にWeb移行を推進中' },
      { taskId: byCode('ORD-011'), ecrsKey: 'E', afterTime: '0', achievement: 'Excel台帳を廃止。基幹に一本化完了。月末の数値突合作業が不要に', date: '2026-03-08', comment: '基幹の受注残レポート機能で代替。追加費用ゼロで実現' }
    ];

    // --- 生産管理プロジェクト: 改善計画3件 ---
    const plnImprovements = [
      { taskId: byCode('PLN-003'), ecrsKey: 'S', content: '生産スケジューラの簡易導入', method: 'Excelマクロベースのスケジューラを構築し機械負荷を可視化', person: '生産管理主任', resources: 'Excel VBA（社内開発）', startDate: '2026-02-15', endDate: '2026-05-15', effectQuantity: '計画立案 120分→40分/週', effectQuality: '計画精度向上、属人化解消', status: 'progress' },
      { taskId: byCode('PLN-006'), ecrsKey: 'S', content: 'タブレット日報入力の導入', method: '現場にタブレット配置しリアルタイムで実績入力', person: '製造部長', resources: 'タブレット3台（15万円）、Googleフォーム', startDate: '2026-04-01', endDate: '2026-06-30', effectQuantity: '進捗確認時間 月40時間→10時間', effectQuality: 'リアルタイム進捗把握', status: 'pending' },
      { taskId: byCode('PLN-008'), ecrsKey: 'S', content: '納期アラートシステムの構築', method: '受注残と進捗を自動照合し遅延リスクをメール通知', person: 'システム担当', resources: 'Excel VBA + Outlook連携（社内開発）', startDate: '2026-04-15', endDate: '2026-06-30', effectQuantity: '遅延検知 1日前→3日前に前倒し', effectQuality: '納期遵守率95%→99%目標', status: 'pending' }
    ];

    // --- 共通の企業情報 ---
    const company = '(株)KK精工';
    const baseMemo = '工業用切削工具・特殊形状工具の製作及び再研削を主業務とする製造業。';

    // --- 5プロジェクトの返却 ---
    return {
      projects: [
        {
          id: 'demo_ord',
          name: '受注プロセス改善',
          company,
          memo: baseMemo + '受注業務（見積依頼受付〜注文確定・納期回答）の効率化と転記ミス削減を目指す。',
          processes: ['受注'],
          currentStep: 5,
          tasks: ord,
          improvements: ordImprovements,
          measurements: ordMeasurements,
          createdAt: '2026-01-10T09:00:00.000Z'
        },
        {
          id: 'demo_pln',
          name: '生産管理プロセス改善',
          company,
          memo: baseMemo + '生産計画立案〜進捗管理・在庫管理の属人化解消とリードタイム短縮を目指す。',
          processes: ['生産管理'],
          currentStep: 4,
          tasks: pln,
          improvements: plnImprovements,
          measurements: [],
          createdAt: '2026-01-20T09:00:00.000Z'
        },
        {
          id: 'demo_mfg',
          name: '製造プロセス改善',
          company,
          memo: baseMemo + 'CNC工具研削・仕上げ加工・日報記録など製造現場の作業効率向上と品質安定化を目指す。',
          processes: ['製造'],
          currentStep: 3,
          tasks: mfg,
          improvements: [],
          measurements: [],
          createdAt: '2026-02-01T09:00:00.000Z'
        },
        {
          id: 'demo_qc',
          name: '品質管理プロセス改善',
          company,
          memo: baseMemo + '検査記録・不適合管理・ISO9001文書管理のデジタル化と検索性向上を目指す。',
          processes: ['品質管理'],
          currentStep: 2,
          tasks: qc,
          improvements: [],
          measurements: [],
          createdAt: '2026-02-15T09:00:00.000Z'
        },
        {
          id: 'demo_shp',
          name: '出荷プロセス改善',
          company,
          memo: baseMemo + '出荷指示〜梱包・送り状作成・配送追跡の業務洗い出しを実施中。',
          processes: ['出荷'],
          currentStep: 1,
          tasks: shp,
          improvements: [],
          measurements: [],
          createdAt: '2026-03-01T09:00:00.000Z'
        }
      ],
      companies: {
        [company]: {
          description: '工業用切削工具（エンドミル・ドリル・リーマ等）および特殊形状工具の製作・再研削を主業務とする製造業。従業員約80名。主に自動車・航空機・金型業界向けに超硬工具を供給。ISO9001認証取得済。'
        }
      },
      processCategories: [...DEFAULT_PROCESS_CATEGORIES],
      currentProjectId: null
    };
  }

  function loadDemoIfEmpty() {
    if (appData.projects && appData.projects.length > 0) return false;
    appData = generateDemoData();
    saveData(appData);
    return true;
  }

  // ==========================================
  // デモデータプレビュー機能
  // ==========================================
  function toggleDemoPreview() {
    if (isDemoPreviewMode) return;
    // 現在のデータをメモリに退避（localStorageには触れない）
    savedDataBeforeDemo = JSON.parse(JSON.stringify(appData));
    isDemoPreviewMode = true;
    // デモデータに切り替え（保存はしない）
    appData = generateDemoData();
    // UI更新
    updateDemoPreviewUI(true);
    navigate('dashboard');
    showToast('デモデータ（(株)KK精工）をプレビュー中です', 'info');
  }

  function exitDemoPreview() {
    if (!isDemoPreviewMode) return;
    // 退避データを復元
    appData = savedDataBeforeDemo;
    savedDataBeforeDemo = null;
    isDemoPreviewMode = false;
    // UI更新
    updateDemoPreviewUI(false);
    navigate('dashboard');
    showToast('元のデータに戻りました', 'success');
  }

  function updateDemoPreviewUI(isDemo) {
    const navDemoPreview = $('#navDemoPreview');
    const navDemoBack = $('#navDemoBack');
    const demoBanner = $('#demoPreviewBanner');
    if (navDemoPreview) navDemoPreview.style.display = isDemo ? 'none' : '';
    if (navDemoBack) navDemoBack.style.display = isDemo ? '' : 'none';
    if (demoBanner) demoBanner.style.display = isDemo ? '' : 'none';
  }

  // ==========================================
  // グローバルAPI（onclick用）
  // ==========================================
  window.BPI = {
    editTask: (id) => openTaskModal(id),
    deleteTask: (id) => deleteTask(id),
    moveTask: (id, dir) => moveTask(id, dir),
    renumberCodes: () => { const p = getCurrentProject(); if (p) renumberTaskCodes(p); },
    openImprovementModal: (taskId, ecrsKey) => openImprovementModal(taskId, ecrsKey),
    updateImpStatus: (taskId, ecrsKey, status) => updateImpStatus(taskId, ecrsKey, status),
    openMeasureModal: (taskId, ecrsKey) => openMeasureModal(taskId, ecrsKey),
    loadDemo: () => { appData = generateDemoData(); saveData(appData); location.reload(); },
    toggleDemoPreview: () => toggleDemoPreview(),
    exitDemoPreview: () => exitDemoPreview(),
    openAIAssist: (step) => openAIAssist(step),
    toggleAISelectAll: (checked) => toggleAISelectAll(checked),
    openCompanyModal: (name) => openCompanyModal(name),
    openNewProjectModal: (company) => openNewProjectModal(company),
    printReport: () => { buildPrintReport(); setTimeout(() => window.print(), 400); },
    exportExcel: () => exportExcel()
  };

  // ==========================================
  // 初期化
  // ==========================================
  function init() {
    initEventListeners();

    // デモデータ読込ボタン
    const btnDemo = $('#btnLoadDemo');
    if (btnDemo) {
      btnDemo.addEventListener('click', () => {
        appData = generateDemoData();
        saveData(appData);
        showToast('デモデータ（(株)KK精工）を読み込みました', 'success');
        navigate('dashboard');
      });
    }

    // 保存されたプロジェクトがある場合
    if (appData.currentProjectId) {
      const proj = getCurrentProject();
      if (proj) {
        $('#currentProjectInfo').style.display = '';
        $('#projectBadge').textContent = `${proj.company} / ${proj.name}`;
      }
    }

    navigate('dashboard');
  }

  // DOM読み込み完了後に初期化
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
