// ===================================================
// CRM 前端 Google Sheets API 整合範例
// ===================================================
// 來源：Alpine.js Store API + Google Apps Script
// https://alpinejs.dev/globals/alpine-store
// https://developers.google.com/apps-script/guides/web
// 實施日期：2024-12-03
// ===================================================

/**
 * Google Sheets API 配置
 * 部署 Apps Script 後，將 Web App URL 填入下方
 */
const GOOGLE_SHEETS_CONFIG = {
    apiUrl: 'PASTE_YOUR_WEB_APP_URL_HERE', // 替換為你的 Web App URL
    timeout: 10000, // API 請求逾時時間（毫秒）
    retryAttempts: 3, // 失敗重試次數
    enableCache: true, // 啟用 localStorage 快取
    cacheExpiry: 300000 // 快取過期時間（5 分鐘）
};

/**
 * Alpine.js Store - Google Sheets API 管理
 * 來源：Alpine.js 官方 Store 文檔
 * https://alpinejs.dev/globals/alpine-store
 */
document.addEventListener('alpine:init', () => {
    Alpine.store('sheets', {
        // 狀態
        loading: false,
        error: null,
        lastSync: null,

        /**
         * 通用 API 請求方法
         * 來源：Fetch API 最佳實踐
         * https://developer.mozilla.org/zh-TW/docs/Web/API/Fetch_API/Using_Fetch
         */
        async request(action, sheet, data = null, id = null) {
            this.loading = true;
            this.error = null;

            try {
                const response = await fetch(GOOGLE_SHEETS_CONFIG.apiUrl, {
                    method: 'POST',
                    mode: 'cors',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        action,
                        sheet,
                        data,
                        id
                    })
                });

                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }

                const result = await response.json();

                if (!result.success) {
                    throw new Error(result.error || 'API request failed');
                }

                this.lastSync = new Date().toISOString();
                this.loading = false;
                return result.data;

            } catch (error) {
                this.error = error.message;
                this.loading = false;
                console.error('API Error:', error);
                throw error;
            }
        },

        /**
         * 讀取數據（含快取）
         * 來源：localStorage 快取策略
         * https://developer.mozilla.org/zh-TW/docs/Web/API/Window/localStorage
         */
        async read(sheet) {
            // 檢查快取
            if (GOOGLE_SHEETS_CONFIG.enableCache) {
                const cached = this.getCache(sheet);
                if (cached) {
                    console.log(`Using cached data for ${sheet}`);
                    return cached;
                }
            }

            // 從 API 讀取
            const data = await this.request('read', sheet);

            // 更新快取
            if (GOOGLE_SHEETS_CONFIG.enableCache) {
                this.setCache(sheet, data);
            }

            return data;
        },

        /**
         * 新增數據
         */
        async create(sheet, data) {
            const result = await this.request('create', sheet, data);
            this.clearCache(sheet); // 清除快取
            return result;
        },

        /**
         * 更新數據
         */
        async update(sheet, id, data) {
            const result = await this.request('update', sheet, data, id);
            this.clearCache(sheet); // 清除快取
            return result;
        },

        /**
         * 刪除數據
         */
        async delete(sheet, id) {
            const result = await this.request('delete', sheet, null, id);
            this.clearCache(sheet); // 清除快取
            return result;
        },

        /**
         * 下載 Excel 備份
         */
        async downloadBackup(sheet = null) {
            return await this.request('backup', sheet);
        },

        /**
         * 快取管理
         */
        getCache(sheet) {
            const key = `crm_cache_${sheet}`;
            const cached = localStorage.getItem(key);

            if (!cached) return null;

            try {
                const { data, timestamp } = JSON.parse(cached);
                const age = Date.now() - timestamp;

                if (age > GOOGLE_SHEETS_CONFIG.cacheExpiry) {
                    this.clearCache(sheet);
                    return null;
                }

                return data;
            } catch (error) {
                console.error('Cache parse error:', error);
                return null;
            }
        },

        setCache(sheet, data) {
            const key = `crm_cache_${sheet}`;
            const value = JSON.stringify({
                data,
                timestamp: Date.now()
            });
            localStorage.setItem(key, value);
        },

        clearCache(sheet) {
            const key = `crm_cache_${sheet}`;
            localStorage.removeItem(key);
        },

        clearAllCache() {
            ['tasks', 'customers', 'cases', 'emails'].forEach(sheet => {
                this.clearCache(sheet);
            });
        }
    });
});

/**
 * 任務管理 Store（使用 Google Sheets）
 * 來源：Alpine.js Reactivity 系統
 * https://alpinejs.dev/advanced/reactivity
 */
document.addEventListener('alpine:init', () => {
    Alpine.store('tasks', {
        items: [],
        loading: false,

        async init() {
            await this.loadTasks();
        },

        async loadTasks() {
            this.loading = true;
            try {
                this.items = await Alpine.store('sheets').read('tasks');
            } catch (error) {
                console.error('Failed to load tasks:', error);
                alert('載入任務失敗：' + error.message);
            } finally {
                this.loading = false;
            }
        },

        async addTask(task) {
            try {
                const result = await Alpine.store('sheets').create('tasks', task);
                await this.loadTasks(); // 重新載入
                return result;
            } catch (error) {
                console.error('Failed to add task:', error);
                alert('新增任務失敗：' + error.message);
                throw error;
            }
        },

        async updateTask(id, updates) {
            try {
                const result = await Alpine.store('sheets').update('tasks', id, updates);
                await this.loadTasks(); // 重新載入
                return result;
            } catch (error) {
                console.error('Failed to update task:', error);
                alert('更新任務失敗：' + error.message);
                throw error;
            }
        },

        async deleteTask(id) {
            try {
                const result = await Alpine.store('sheets').delete('tasks', id);
                await this.loadTasks(); // 重新載入
                return result;
            } catch (error) {
                console.error('Failed to delete task:', error);
                alert('刪除任務失敗：' + error.message);
                throw error;
            }
        },

        // 本地篩選（不需要 API 調用）
        get todayTasks() {
            const today = new Date().toDateString();
            return this.items.filter(task => {
                const dueDate = new Date(task['到期日']);
                return dueDate.toDateString() === today;
            });
        },

        get overdueTasks() {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            return this.items.filter(task => {
                const dueDate = new Date(task['到期日']);
                return dueDate < today && task['狀態'] !== 'completed';
            });
        },

        get completedTasks() {
            return this.items.filter(task => task['狀態'] === 'completed');
        }
    });
});

/**
 * 使用範例：HTML 整合
 *
 * <div x-data x-init="$store.tasks.init()">
 *     <!-- 載入狀態 -->
 *     <div x-show="$store.sheets.loading" class="text-blue-600">
 *         載入中...
 *     </div>
 *
 *     <!-- 錯誤訊息 -->
 *     <div x-show="$store.sheets.error" class="text-red-600" x-text="$store.sheets.error"></div>
 *
 *     <!-- 任務列表 -->
 *     <template x-for="task in $store.tasks.items" :key="task.ID">
 *         <div>
 *             <h3 x-text="task['標題']"></h3>
 *             <p x-text="task['描述']"></p>
 *             <button @click="$store.tasks.deleteTask(task.ID)">刪除</button>
 *         </div>
 *     </template>
 *
 *     <!-- 新增任務表單 -->
 *     <form @submit.prevent="
 *         $store.tasks.addTask({
 *             '標題': $refs.title.value,
 *             '描述': $refs.desc.value,
 *             '狀態': 'pending',
 *             '到期日': $refs.date.value
 *         })
 *     ">
 *         <input x-ref="title" placeholder="標題" required>
 *         <textarea x-ref="desc" placeholder="描述"></textarea>
 *         <input x-ref="date" type="date" required>
 *         <button type="submit">新增任務</button>
 *     </form>
 *
 *     <!-- Excel 備份按鈕 -->
 *     <button @click="$store.sheets.downloadBackup('tasks')">
 *         下載 Excel 備份
 *     </button>
 * </div>
 */
