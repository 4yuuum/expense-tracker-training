/**
 * 支出記録アプリ - メインスクリプト
 * localStorageを使用してデータを永続化
 */

// ========================================
// 定数定義
// ========================================

// localStorageのキー
const STORAGE_KEYS = {
    EXPENSES: 'expenseTracker_expenses',
    CATEGORIES: 'expenseTracker_categories',
    SETTINGS: 'expenseTracker_settings'
};

// デフォルトカテゴリ
const DEFAULT_CATEGORIES = ['食費', '交通費', '娯楽', '日用品', '光熱費', 'その他'];

// 設定のデフォルト値
const DEFAULT_SETTINGS = {
    currency: '¥',
    dateFormat: 'YYYY-MM-DD',
    defaultCategory: 'その他'
};

// ========================================
// データ管理ユーティリティ
// ========================================

/**
 * 支出データを読み込む
 * @returns {Array} 支出データの配列
 */
function loadExpenses() {
    try {
        const data = localStorage.getItem(STORAGE_KEYS.EXPENSES);
        return data ? JSON.parse(data) : [];
    } catch (error) {
        console.error('支出データの読み込みに失敗しました:', error);
        return [];
    }
}

/**
 * 支出データを保存する
 * @param {Array} expenses - 支出データの配列
 */
function saveExpenses(expenses) {
    try {
        localStorage.setItem(STORAGE_KEYS.EXPENSES, JSON.stringify(expenses));
    } catch (error) {
        console.error('支出データの保存に失敗しました:', error);
        alert('データの保存に失敗しました。ブラウザのストレージ容量を確認してください。');
    }
}

/**
 * カテゴリリストを読み込む
 * @returns {Array} カテゴリの配列
 */
function loadCategories() {
    try {
        const data = localStorage.getItem(STORAGE_KEYS.CATEGORIES);
        return data ? JSON.parse(data) : DEFAULT_CATEGORIES;
    } catch (error) {
        console.error('カテゴリデータの読み込みに失敗しました:', error);
        return DEFAULT_CATEGORIES;
    }
}

/**
 * カテゴリリストを保存する
 * @param {Array} categories - カテゴリの配列
 */
function saveCategories(categories) {
    try {
        localStorage.setItem(STORAGE_KEYS.CATEGORIES, JSON.stringify(categories));
    } catch (error) {
        console.error('カテゴリデータの保存に失敗しました:', error);
    }
}

/**
 * 設定を読み込む
 * @returns {Object} 設定オブジェクト
 */
function loadSettings() {
    try {
        const data = localStorage.getItem(STORAGE_KEYS.SETTINGS);
        return data ? JSON.parse(data) : DEFAULT_SETTINGS;
    } catch (error) {
        console.error('設定データの読み込みに失敗しました:', error);
        return DEFAULT_SETTINGS;
    }
}

/**
 * 初期化処理（初回起動時にデフォルト値を設定）
 */
function initializeData() {
    // カテゴリが存在しない場合はデフォルトカテゴリを設定
    if (!localStorage.getItem(STORAGE_KEYS.CATEGORIES)) {
        saveCategories(DEFAULT_CATEGORIES);
    }
    // 設定が存在しない場合はデフォルト設定を設定
    if (!localStorage.getItem(STORAGE_KEYS.SETTINGS)) {
        localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(DEFAULT_SETTINGS));
    }
}

// ========================================
// 支出管理機能
// ========================================

/**
 * 一意のIDを生成する
 * @returns {string} 一意のID
 */
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

/**
 * 支出を追加する
 * @param {Object} expenseData - 支出データ（日付、カテゴリ、金額、メモ）
 * @returns {boolean} 成功した場合true
 */
function addExpense(expenseData) {
    const expenses = loadExpenses();
    const newExpense = {
        id: generateId(),
        date: expenseData.date,
        category: expenseData.category,
        amount: parseInt(expenseData.amount, 10),
        memo: expenseData.memo || '',
        createdAt: Date.now()
    };
    expenses.push(newExpense);
    saveExpenses(expenses);
    return true;
}

/**
 * 支出を更新する
 * @param {string} id - 支出のID
 * @param {Object} updatedData - 更新するデータ
 * @returns {boolean} 成功した場合true
 */
function updateExpense(id, updatedData) {
    const expenses = loadExpenses();
    const index = expenses.findIndex(e => e.id === id);
    if (index === -1) {
        return false;
    }
    expenses[index] = {
        ...expenses[index],
        ...updatedData,
        amount: parseInt(updatedData.amount, 10),
        updatedAt: Date.now()
    };
    saveExpenses(expenses);
    return true;
}

/**
 * 支出を削除する
 * @param {string} id - 支出のID
 * @returns {boolean} 成功した場合true
 */
function deleteExpense(id) {
    const expenses = loadExpenses();
    const filtered = expenses.filter(e => e.id !== id);
    saveExpenses(filtered);
    return true;
}

/**
 * IDで支出を取得する
 * @param {string} id - 支出のID
 * @returns {Object|null} 支出オブジェクト、見つからない場合はnull
 */
function getExpenseById(id) {
    const expenses = loadExpenses();
    return expenses.find(e => e.id === id) || null;
}

// ========================================
// フィルタリング・ソート機能
// ========================================

/**
 * 支出データをフィルタリングする
 * @param {Array} expenses - 支出データの配列
 * @param {Object} filters - フィルタ条件
 * @returns {Array} フィルタリングされた支出データ
 */
function filterExpenses(expenses, filters) {
    return expenses.filter(expense => {
        // 日付範囲フィルタ
        if (filters.dateFrom && expense.date < filters.dateFrom) {
            return false;
        }
        if (filters.dateTo && expense.date > filters.dateTo) {
            return false;
        }
        // カテゴリフィルタ
        if (filters.category && expense.category !== filters.category) {
            return false;
        }
        // 金額範囲フィルタ
        if (filters.amountMin !== null && expense.amount < filters.amountMin) {
            return false;
        }
        if (filters.amountMax !== null && expense.amount > filters.amountMax) {
            return false;
        }
        // メモ検索フィルタ
        if (filters.searchMemo && !expense.memo.toLowerCase().includes(filters.searchMemo.toLowerCase())) {
            return false;
        }
        return true;
    });
}

/**
 * 支出データをソートする
 * @param {Array} expenses - 支出データの配列
 * @param {string} sortBy - ソート方法（'date-desc', 'date-asc', 'amount-desc', 'amount-asc', 'category-asc'）
 * @returns {Array} ソートされた支出データ
 */
function sortExpenses(expenses, sortBy) {
    const sorted = [...expenses];
    switch (sortBy) {
        case 'date-desc':
            return sorted.sort((a, b) => new Date(b.date) - new Date(a.date));
        case 'date-asc':
            return sorted.sort((a, b) => new Date(a.date) - new Date(b.date));
        case 'amount-desc':
            return sorted.sort((a, b) => b.amount - a.amount);
        case 'amount-asc':
            return sorted.sort((a, b) => a.amount - b.amount);
        case 'category-asc':
            return sorted.sort((a, b) => a.category.localeCompare(b.category, 'ja'));
        default:
            return sorted;
    }
}

// ========================================
// 統計計算機能
// ========================================

/**
 * 金額をフォーマットする
 * @param {number} amount - 金額
 * @returns {string} フォーマットされた金額文字列
 */
function formatAmount(amount) {
    return `¥${amount.toLocaleString('ja-JP')}`;
}

/**
 * 日付をフォーマットする（YYYY/MM/DD形式）
 * @param {string} dateString - 日付文字列（YYYY-MM-DD形式）
 * @returns {string} フォーマットされた日付文字列
 */
function formatDate(dateString) {
    const date = new Date(dateString);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
}

/**
 * 今月の合計支出を計算する
 * @param {Array} expenses - 支出データの配列
 * @returns {number} 合計金額
 */
function calculateMonthlyTotal(expenses) {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    return expenses
        .filter(expense => {
            const expenseDate = new Date(expense.date);
            return expenseDate.getMonth() === currentMonth && 
                   expenseDate.getFullYear() === currentYear;
        })
        .reduce((sum, expense) => sum + expense.amount, 0);
}

/**
 * 今年の合計支出を計算する
 * @param {Array} expenses - 支出データの配列
 * @returns {number} 合計金額
 */
function calculateYearlyTotal(expenses) {
    const currentYear = new Date().getFullYear();
    
    return expenses
        .filter(expense => {
            const expenseDate = new Date(expense.date);
            return expenseDate.getFullYear() === currentYear;
        })
        .reduce((sum, expense) => sum + expense.amount, 0);
}

/**
 * 平均日次支出を計算する
 * @param {Array} expenses - 支出データの配列
 * @returns {number} 平均金額
 */
function calculateAverageDaily(expenses) {
    if (expenses.length === 0) return 0;
    
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    const firstDayOfMonth = new Date(currentYear, currentMonth, 1);
    const daysPassed = Math.ceil((now - firstDayOfMonth) / (1000 * 60 * 60 * 24));
    
    const monthlyTotal = calculateMonthlyTotal(expenses);
    return daysPassed > 0 ? Math.round(monthlyTotal / daysPassed) : 0;
}

/**
 * カテゴリ別の集計を計算する
 * @param {Array} expenses - 支出データの配列
 * @returns {Array} カテゴリ別集計データ
 */
function calculateCategoryStats(expenses) {
    const categoryMap = {};
    let total = 0;
    
    // カテゴリ別に合計を計算
    expenses.forEach(expense => {
        if (!categoryMap[expense.category]) {
            categoryMap[expense.category] = 0;
        }
        categoryMap[expense.category] += expense.amount;
        total += expense.amount;
    });
    
    // 配列に変換してソート
    const stats = Object.entries(categoryMap)
        .map(([category, amount]) => ({
            category,
            amount,
            percentage: total > 0 ? Math.round((amount / total) * 100) : 0
        }))
        .sort((a, b) => b.amount - a.amount);
    
    return stats;
}

// ========================================
// DOM操作・UI更新機能
// ========================================

/**
 * カテゴリ選択肢を設定する
 */
function populateCategories() {
    const categories = loadCategories();
    const categorySelect = document.getElementById('expenseCategory');
    const categoryFilter = document.getElementById('categoryFilter');
    
    // フォームのカテゴリ選択肢を更新
    categorySelect.innerHTML = '<option value="">選択してください</option>';
    categories.forEach(category => {
        const option = document.createElement('option');
        option.value = category;
        option.textContent = category;
        categorySelect.appendChild(option);
    });
    
    // フィルタのカテゴリ選択肢を更新
    categoryFilter.innerHTML = '<option value="">すべて</option>';
    categories.forEach(category => {
        const option = document.createElement('option');
        option.value = category;
        option.textContent = category;
        categoryFilter.appendChild(option);
    });
}

/**
 * 統計情報を更新する
 */
function updateStatistics() {
    const expenses = loadExpenses();
    const monthlyTotal = calculateMonthlyTotal(expenses);
    const yearlyTotal = calculateYearlyTotal(expenses);
    const averageDaily = calculateAverageDaily(expenses);
    
    document.getElementById('monthlyTotal').textContent = formatAmount(monthlyTotal);
    document.getElementById('yearlyTotal').textContent = formatAmount(yearlyTotal);
    document.getElementById('averageDaily').textContent = formatAmount(averageDaily);
}

/**
 * 支出一覧を表示する
 */
function renderExpenseList() {
    const expenses = loadExpenses();
    
    // フィルタ条件を取得
    const filters = {
        dateFrom: document.getElementById('dateFrom').value || null,
        dateTo: document.getElementById('dateTo').value || null,
        category: document.getElementById('categoryFilter').value || null,
        amountMin: document.getElementById('amountMin').value ? 
                   parseInt(document.getElementById('amountMin').value, 10) : null,
        amountMax: document.getElementById('amountMax').value ? 
                   parseInt(document.getElementById('amountMax').value, 10) : null,
        searchMemo: document.getElementById('searchMemo').value || null
    };
    
    // ソート方法を取得
    const sortBy = document.getElementById('sortBy').value;
    
    // フィルタリングとソートを実行
    let filteredExpenses = filterExpenses(expenses, filters);
    filteredExpenses = sortExpenses(filteredExpenses, sortBy);
    
    // テーブルボディを取得
    const tableBody = document.getElementById('expenseTableBody');
    const emptyMessage = document.getElementById('emptyMessage');
    
    // 空の場合はメッセージを表示
    if (filteredExpenses.length === 0) {
        tableBody.innerHTML = '';
        emptyMessage.style.display = 'block';
        document.getElementById('expenseCount').textContent = '0件の支出';
        document.getElementById('filteredTotal').textContent = '合計: ¥0';
        return;
    }
    
    emptyMessage.style.display = 'none';
    
    // テーブル行を生成
    tableBody.innerHTML = filteredExpenses.map(expense => {
        return `
            <tr>
                <td>${formatDate(expense.date)}</td>
                <td><span class="category-badge">${expense.category}</span></td>
                <td class="amount-cell">${formatAmount(expense.amount)}</td>
                <td class="memo-cell" title="${expense.memo}">${expense.memo || '-'}</td>
                <td>
                    <div class="action-buttons">
                        <button class="btn btn--primary btn--small" onclick="editExpense('${expense.id}')" aria-label="編集">
                            編集
                        </button>
                        <button class="btn btn--danger btn--small" onclick="confirmDeleteExpense('${expense.id}')" aria-label="削除">
                            削除
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
    
    // 件数と合計を更新
    const total = filteredExpenses.reduce((sum, e) => sum + e.amount, 0);
    document.getElementById('expenseCount').textContent = `${filteredExpenses.length}件の支出`;
    document.getElementById('filteredTotal').textContent = `合計: ${formatAmount(total)}`;
}

/**
 * カテゴリ別統計を表示する
 */
function renderCategoryStats() {
    const expenses = loadExpenses();
    
    // フィルタ条件を取得（支出一覧と同じ条件を使用）
    const filters = {
        dateFrom: document.getElementById('dateFrom').value || null,
        dateTo: document.getElementById('dateTo').value || null,
        category: null, // カテゴリ別統計ではカテゴリフィルタは適用しない
        amountMin: document.getElementById('amountMin').value ? 
                   parseInt(document.getElementById('amountMin').value, 10) : null,
        amountMax: document.getElementById('amountMax').value ? 
                   parseInt(document.getElementById('amountMax').value, 10) : null,
        searchMemo: document.getElementById('searchMemo').value || null
    };
    
    let filteredExpenses = filterExpenses(expenses, filters);
    const stats = calculateCategoryStats(filteredExpenses);
    
    const categoryStatsContainer = document.getElementById('categoryStats');
    
    if (stats.length === 0) {
        categoryStatsContainer.innerHTML = '<p style="text-align: center; color: var(--color-text-light);">データがありません</p>';
        return;
    }
    
    categoryStatsContainer.innerHTML = stats.map(stat => {
        return `
            <div class="category-stat-item">
                <div class="category-stat-item__name">${stat.category}</div>
                <div class="category-stat-item__amount">${formatAmount(stat.amount)}</div>
                <div class="category-stat-item__percentage">${stat.percentage}%</div>
            </div>
        `;
    }).join('');
}

/**
 * 画面全体を更新する
 */
function refreshUI() {
    updateStatistics();
    renderExpenseList();
    renderCategoryStats();
}

// ========================================
// モーダル管理機能
// ========================================

let editingExpenseId = null; // 編集中の支出ID

/**
 * モーダルを開く
 * @param {string} modalId - モーダルのID
 */
function openModal(modalId) {
    const modal = document.getElementById(modalId);
    modal.setAttribute('aria-hidden', 'false');
    modal.style.display = 'flex';
}

/**
 * モーダルを閉じる
 * @param {string} modalId - モーダルのID
 */
function closeModal(modalId) {
    const modal = document.getElementById(modalId);
    modal.setAttribute('aria-hidden', 'true');
    modal.style.display = 'none';
}

/**
 * 支出追加モーダルを開く
 */
function openAddExpenseModal() {
    editingExpenseId = null;
    document.getElementById('modalTitle').textContent = '支出を追加';
    document.getElementById('expenseForm').reset();
    document.getElementById('expenseDate').valueAsDate = new Date(); // 今日の日付を設定
    document.getElementById('formError').classList.remove('show');
    openModal('expenseModal');
}

/**
 * 支出編集モーダルを開く
 * @param {string} id - 支出のID
 */
function editExpense(id) {
    const expense = getExpenseById(id);
    if (!expense) {
        alert('支出が見つかりませんでした。');
        return;
    }
    
    editingExpenseId = id;
    document.getElementById('modalTitle').textContent = '支出を編集';
    document.getElementById('expenseDate').value = expense.date;
    document.getElementById('expenseCategory').value = expense.category;
    document.getElementById('expenseAmount').value = expense.amount;
    document.getElementById('expenseMemo').value = expense.memo || '';
    document.getElementById('formError').classList.remove('show');
    openModal('expenseModal');
}

/**
 * 支出削除確認モーダルを開く
 * @param {string} id - 支出のID
 */
function confirmDeleteExpense(id) {
    const expense = getExpenseById(id);
    if (!expense) {
        alert('支出が見つかりませんでした。');
        return;
    }
    
    document.getElementById('deleteInfo').textContent = 
        `${formatDate(expense.date)} - ${expense.category} - ${formatAmount(expense.amount)}`;
    
    // 削除確認ボタンにIDを保存
    document.getElementById('confirmDeleteBtn').dataset.expenseId = id;
    openModal('deleteModal');
}

/**
 * 支出を削除する（確認後）
 */
function handleDeleteExpense() {
    const id = document.getElementById('confirmDeleteBtn').dataset.expenseId;
    if (id && deleteExpense(id)) {
        closeModal('deleteModal');
        refreshUI();
    } else {
        alert('削除に失敗しました。');
    }
}

// ========================================
// フォーム処理
// ========================================

/**
 * フォームのバリデーション
 * @param {Object} formData - フォームデータ
 * @returns {Object} { valid: boolean, errors: Array<string> }
 */
function validateForm(formData) {
    const errors = [];
    
    if (!formData.date) {
        errors.push('日付を入力してください。');
    }
    
    if (!formData.category) {
        errors.push('カテゴリを選択してください。');
    }
    
    if (!formData.amount || formData.amount <= 0) {
        errors.push('金額を正の数値で入力してください。');
    }
    
    if (formData.memo && formData.memo.length > 200) {
        errors.push('メモは200文字以内で入力してください。');
    }
    
    return {
        valid: errors.length === 0,
        errors
    };
}

/**
 * フォーム送信処理
 */
function handleFormSubmit(event) {
    event.preventDefault();
    
    // フォームデータを取得
    const formData = {
        date: document.getElementById('expenseDate').value,
        category: document.getElementById('expenseCategory').value,
        amount: document.getElementById('expenseAmount').value,
        memo: document.getElementById('expenseMemo').value.trim()
    };
    
    // バリデーション
    const validation = validateForm(formData);
    const errorElement = document.getElementById('formError');
    
    if (!validation.valid) {
        errorElement.textContent = validation.errors.join('\n');
        errorElement.classList.add('show');
        return;
    }
    
    errorElement.classList.remove('show');
    
    // 追加または更新
    let success = false;
    if (editingExpenseId) {
        success = updateExpense(editingExpenseId, formData);
    } else {
        success = addExpense(formData);
    }
    
    if (success) {
        closeModal('expenseModal');
        refreshUI();
    } else {
        alert('保存に失敗しました。');
    }
}

// ========================================
// データエクスポート・インポート機能
// ========================================

/**
 * データをエクスポートする（JSON形式）
 */
function exportData() {
    const expenses = loadExpenses();
    const categories = loadCategories();
    const settings = loadSettings();
    
    const data = {
        expenses,
        categories,
        settings,
        exportDate: new Date().toISOString()
    };
    
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `expense-tracker-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

/**
 * データをインポートする（JSON形式）
 * @param {File} file - インポートするJSONファイル
 */
function importData(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = JSON.parse(e.target.result);
            
            if (!confirm('インポートすると既存のデータが上書きされます。よろしいですか？')) {
                return;
            }
            
            // データを保存
            if (data.expenses) {
                saveExpenses(data.expenses);
            }
            if (data.categories) {
                saveCategories(data.categories);
                populateCategories();
            }
            if (data.settings) {
                localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(data.settings));
            }
            
            alert('データのインポートが完了しました。');
            refreshUI();
        } catch (error) {
            console.error('インポートエラー:', error);
            alert('ファイルの読み込みに失敗しました。正しいJSON形式のファイルを選択してください。');
        }
    };
    reader.readAsText(file);
}

// ========================================
// イベントリスナー設定
// ========================================

/**
 * アプリケーションの初期化
 */
function init() {
    // データの初期化
    initializeData();
    
    // カテゴリ選択肢を設定
    populateCategories();
    
    // UIを更新
    refreshUI();
    
    // イベントリスナーを設定
    
    // 支出追加ボタン
    document.getElementById('addExpenseBtn').addEventListener('click', openAddExpenseModal);
    
    // モーダルの閉じるボタン
    document.getElementById('closeModalBtn').addEventListener('click', () => {
        closeModal('expenseModal');
    });
    
    document.getElementById('cancelBtn').addEventListener('click', () => {
        closeModal('expenseModal');
    });
    
    // モーダルのオーバーレイクリックで閉じる
    document.querySelectorAll('.modal__overlay').forEach(overlay => {
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) {
                const modal = overlay.closest('.modal');
                closeModal(modal.id);
            }
        });
    });
    
    // フォーム送信
    document.getElementById('expenseForm').addEventListener('submit', handleFormSubmit);
    
    // 削除確認
    document.getElementById('confirmDeleteBtn').addEventListener('click', handleDeleteExpense);
    document.getElementById('cancelDeleteBtn').addEventListener('click', () => {
        closeModal('deleteModal');
    });
    
    // フィルタリング・ソートの変更時に一覧を更新
    const filterInputs = [
        'dateFrom', 'dateTo', 'categoryFilter', 
        'amountMin', 'amountMax', 'searchMemo', 'sortBy'
    ];
    filterInputs.forEach(id => {
        document.getElementById(id).addEventListener('change', refreshUI);
        document.getElementById(id).addEventListener('input', refreshUI);
    });
    
    // フィルタクリアボタン
    document.getElementById('clearFiltersBtn').addEventListener('click', () => {
        document.getElementById('dateFrom').value = '';
        document.getElementById('dateTo').value = '';
        document.getElementById('categoryFilter').value = '';
        document.getElementById('amountMin').value = '';
        document.getElementById('amountMax').value = '';
        document.getElementById('searchMemo').value = '';
        document.getElementById('sortBy').value = 'date-desc';
        refreshUI();
    });
}

// DOMContentLoaded時に初期化
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}


// 追加分ここから
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyHe24aV44cu7xG7Ch27qOP9LKE3uNS9SIoG-7grt5fCaJytOwt3s7K3BGtSGnsIWKx/exec";  // GAS Web App URL
const SECRET = "abcdabcd1234"; // GAS と同じ値

async function syncToSheet() {
    const data = {};
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        let value = localStorage.getItem(key);
        try { value = JSON.parse(value); } catch(e) {}
        data[key] = value;
    }

    try {
        const result = await fetch(SCRIPT_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain" },  // ← プリフライトを防ぐ
            body: JSON.stringify({ secret: SECRET, data: data })
        });

        const json = await result.json();

        if (json.ok) {
            alert("スプレッドシートに同期しました！");
        } else {
            alert("エラー: " + json.error);
        }
    } catch (err) {
        console.error(err);
        alert("通信エラーが発生しました。");
    }
}

// ボタンにイベントを設定（DOM はすでに読み込まれている想定）
document.getElementById("syncButton").addEventListener("click", syncToSheet);