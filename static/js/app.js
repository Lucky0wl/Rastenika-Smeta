document.addEventListener('DOMContentLoaded', () => {
    // Sidebar Toggle
    const toggleBtn = document.getElementById('toggle-sidebar');
    const sidebar = document.querySelector('.sidebar');

    // Загружаем состояние из localStorage
    const sidebarCollapsed = localStorage.getItem('sidebarCollapsed') === 'true';
    if (sidebarCollapsed) {
        sidebar.classList.add('collapsed');
    }

    toggleBtn.addEventListener('click', () => {
        sidebar.classList.toggle('collapsed');
        localStorage.setItem('sidebarCollapsed', sidebar.classList.contains('collapsed'));
    });
    // Transliteration mapping (Cyrillic to Latin)
    const translit = {
        'А': 'A', 'а': 'a', 'Б': 'B', 'б': 'b', 'В': 'V', 'в': 'v', 'Г': 'G', 'г': 'g',
        'Д': 'D', 'д': 'd', 'Е': 'E', 'е': 'e', 'Ё': 'Yo', 'ё': 'yo', 'Ж': 'Zh', 'ж': 'zh',
        'З': 'Z', 'з': 'z', 'И': 'I', 'и': 'i', 'Й': 'Y', 'й': 'y', 'К': 'K', 'к': 'k',
        'Л': 'L', 'л': 'l', 'М': 'M', 'м': 'm', 'Н': 'N', 'н': 'n', 'О': 'O', 'о': 'o',
        'П': 'P', 'п': 'p', 'Р': 'R', 'р': 'r', 'С': 'S', 'с': 's', 'Т': 'T', 'т': 't',
        'У': 'U', 'у': 'u', 'Ф': 'F', 'ф': 'f', 'Х': 'H', 'х': 'h', 'Ц': 'Ts', 'ц': 'ts',
        'Ч': 'Ch', 'ч': 'ch', 'Ш': 'Sh', 'ш': 'sh', 'Щ': 'Sch', 'щ': 'sch',
        'Ъ': '', 'ъ': '', 'Ы': 'Y', 'ы': 'y', 'Ь': '', 'ь': '', 'Э': 'E', 'э': 'e',
        'Ю': 'Yu', 'ю': 'yu', 'Я': 'Ya', 'я': 'ya'
    };

    // Convert Cyrillic to Latin
    function transliterateString(str) {
        return str.split('').map(char => translit[char] || char).join('');
    }

    // Normalize string for search (handle both Cyrillic and Latin)
    function normalizeForSearch(str) {
        return transliterateString(str).toLowerCase();
    }

    // State
    let priceList = [];
    let estimate = [];
    let config = {
        taxRate: 0,
        plantingMethod: 'percent', // 'percent' or 'fixed'
        plantingValue: 35 // 35% default
    };

    // Database view state
    let dbDisplayed = [];
    let dbCurrentPage = 0;
    let dbCurrentSort = null;
    let dbCurrentFilter = '';

    // localStorage persistence
    function saveState() {
        try {
            localStorage.setItem('rastenika_estimate', JSON.stringify(estimate));
            localStorage.setItem('rastenika_config', JSON.stringify(config));
        } catch (e) { /* ignore quota errors */ }
    }

    function loadState() {
        try {
            const savedEstimate = localStorage.getItem('rastenika_estimate');
            const savedConfig = localStorage.getItem('rastenika_config');
            if (savedEstimate) estimate = JSON.parse(savedEstimate);
            if (savedConfig) {
                const c = JSON.parse(savedConfig);
                config = { ...config, ...c };
                // Restore settings inputs
                const taxRateInput = document.getElementById('tax-rate');
                const plantingPercentInput = document.getElementById('planting-percent');
                if (taxRateInput && config.taxRate) taxRateInput.value = config.taxRate;
                if (plantingPercentInput && config.plantingValue) plantingPercentInput.value = config.plantingValue;
            }
        } catch (e) { /* ignore parse errors */ }
    }
    loadState();

    // DOM Elements — defined below, renderEstimate() called after autoLoadPlants()
    const dropZone = document.getElementById('drop-zone');
    const dbStatus = document.getElementById('db-status');
    const fileInput = document.getElementById('file-input');
    const plantSearch = document.getElementById('plant-search');
    const searchResults = document.getElementById('search-results');
    const estimateTable = document.querySelector('#estimate-table tbody');
    const emptyState = document.getElementById('empty-state');
    const generatePdfBtn = document.getElementById('generate-pdf-btn');
    const generateXlsxBtn = document.getElementById('generate-xlsx-btn');
    const dbSearch = document.getElementById('db-search');
    const dbTableBody = document.querySelector('#db-table tbody');

    // Preview Modal Elements
    const previewBtn = document.getElementById('preview-btn');
    const previewModal = document.getElementById('preview-modal');
    const closeModalBtn = document.getElementById('close-modal');
    const previewIframe = document.getElementById('preview-iframe');

    // Summary Elements
    const sumMaterialsEl = document.getElementById('sum-materials');
    const sumLaborEl = document.getElementById('sum-labor');
    const sumTaxEl = document.getElementById('sum-tax');
    const grandTotalEl = document.getElementById('grand-total');
    const displayTaxEl = document.getElementById('display-tax');

    // Navigation
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const view = btn.dataset.view;
            document.querySelectorAll('.view').forEach(v => v.classList.add('hidden'));
            document.getElementById(`view-${view}`).classList.remove('hidden');
            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            if (view === 'database') {
                renderDatabase();
            }
        });
    });

    // File Upload
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => handleFileUpload(e.target.files[0]));

    // Auto-load plants from server on page load
    async function autoLoadPlants() {
        try {
            const response = await fetch('/get-plants');
            const plants = await response.json();
            if (plants && plants.length > 0) {
                priceList = plants;
                plantSearch.disabled = false;
                dbStatus.innerHTML = `<i data-lucide="check-circle" style="color: var(--primary)"></i> <span>База загружена (${plants.length} позиций)</span>`;
                lucide.createIcons();
                renderDatabase();
            }
        } catch (err) {
            // Silent fail — server might not have data yet
        }
        // Восстанавливаем сохранённую смету после загрузки страницы
        if (estimate.length > 0) {
            renderEstimate();
        }
    }
    autoLoadPlants();

    async function handleFileUpload(file) {
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        showToast('Загрузка прайса...');

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });
            const data = await response.json();

            if (data.error) {
                showToast(data.error, 'error');
            } else {
                priceList = data.items;
                plantSearch.disabled = false;
                showToast('Прайс успешно загружен!');
                dbStatus.innerHTML = `<i data-lucide="check-circle" style="color: var(--primary)"></i> <span>${file.name}</span>`;
                lucide.createIcons();
                renderDatabase();
            }
        } catch (err) {
            showToast('Ошибка при загрузке файла', 'error');
        }
    }

    // Database View Rendering with Pagination and Sorting
    function renderDatabase(filter = '', reset = true) {
        if (reset) {
            dbCurrentPage = 0;
            dbCurrentFilter = normalizeForSearch(filter);
        }

        // Filter by search and parameters
        let filtered = priceList.filter(item => {
            const name = normalizeForSearch(item.name || '');
            const params = normalizeForSearch(item.parameters || '');
            const query = dbCurrentFilter;
            return name.includes(query) || params.includes(query);
        });

        // Apply sorting
        if (dbCurrentSort === 'name') {
            filtered.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
        } else if (dbCurrentSort === 'name-desc') {
            filtered.sort((a, b) => (b.name || '').localeCompare(a.name || ''));
        } else if (dbCurrentSort === 'price') {
            filtered.sort((a, b) => (a.price || 0) - (b.price || 0));
        } else if (dbCurrentSort === 'price-desc') {
            filtered.sort((a, b) => (b.price || 0) - (a.price || 0));
        }

        dbDisplayed = filtered;
        const itemsPerPage = 100;
        const start = dbCurrentPage * itemsPerPage;
        const end = start + itemsPerPage;
        const pageItems = filtered.slice(start, end);
        const hasMore = filtered.length > end;

        // Update counter
        const shown = Math.min(end, filtered.length);
        document.getElementById('db-count-shown').innerText = shown;
        document.getElementById('db-count-total').innerText = filtered.length;

        // Render table
        dbTableBody.innerHTML = pageItems.map((item) => `
            <tr>
                <td>${item.name || '-'}</td>
                <td>-</td>
                <td>${item.parameters || '-'}</td>
                <td>${item.price || 0} ₽</td>
                <td>
                    <button class="add-db-btn btn-primary" data-name="${item.name}" style="padding: 6px 12px; font-size: 0.8rem;">
                        <i data-lucide="plus"></i> Добавить
                    </button>
                </td>
            </tr>
        `).join('');

        // Show/hide "Load More" button
        const loadMoreBtn = document.getElementById('db-load-more');
        if (hasMore) {
            loadMoreBtn.style.display = 'block';
        } else {
            loadMoreBtn.style.display = 'none';
        }

        lucide.createIcons();

        // Attach event listeners
        document.querySelectorAll('.add-db-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const itemName = btn.dataset.name;
                const item = priceList.find(p => p.name === itemName);
                if (item) addToEstimate(item);
            });
        });
    }

    dbSearch.addEventListener('input', (e) => {
        renderDatabase(e.target.value);
    });

    // Sorting buttons
    document.getElementById('db-sort-name').addEventListener('click', () => {
        if (dbCurrentSort === 'name') {
            dbCurrentSort = 'name-desc';
        } else {
            dbCurrentSort = 'name';
        }
        renderDatabase(dbCurrentFilter, false);
    });

    document.getElementById('db-sort-price').addEventListener('click', () => {
        if (dbCurrentSort === 'price') {
            dbCurrentSort = 'price-desc';
        } else {
            dbCurrentSort = 'price';
        }
        renderDatabase(dbCurrentFilter, false);
    });

    // Load More button
    document.getElementById('db-load-more').addEventListener('click', () => {
        dbCurrentPage++;
        renderDatabase(dbCurrentFilter, false);
    });

    // Search Logic (Constructor View)
    plantSearch.addEventListener('input', (e) => {
        const query = normalizeForSearch(e.target.value);
        if (query.length < 1) {
            searchResults.classList.add('hidden');
            return;
        }

        const filtered = priceList.filter(item => {
            const name = normalizeForSearch(item.name || '');
            const params = normalizeForSearch(item.parameters || '');
            return name.includes(query) || params.includes(query);
        }).slice(0, 10);

        renderSearchResults(filtered);
    });

    function renderSearchResults(results) {
        if (results.length === 0) {
            searchResults.classList.add('hidden');
            return;
        }

        searchResults.innerHTML = results.map(item => `
            <div class="search-item" data-id="${item.id || Math.random()}">
                <div class="name">${item.name}</div>
                <div style="font-size: 0.85rem; color: #64748B;">${item.parameters || '-'}</div>
                <div class="price">${item.price || 0} ₽</div>
            </div>
        `).join('');

        searchResults.classList.remove('hidden');

        document.querySelectorAll('.search-item').forEach((row, index) => {
            row.addEventListener('click', () => {
                addToEstimate(results[index]);
                searchResults.classList.add('hidden');
                plantSearch.value = '';
            });
        });
    }

    // Estimate Management
    function isMaterial(name) {
        const n = name.toLowerCase();
        return n.includes('кора') || n.includes('торф') || n.includes('грунт') ||
            n.includes('удобрение') || n.includes('камень') || n.includes('мульча') ||
            n.includes('доставка') || n.includes('песок') || n.includes('щебень') ||
            n.includes('корневин') || n.includes('химия');
    }

    function addToEstimate(item) {
        const type = isMaterial(item.name || '') ? 'material' : 'plant';
        const newItem = {
            id: Date.now(),
            type: type,
            name: item.name,
            parameters: item.parameters || '',
            price: parseFloat(item.price) || 0,
            quantity: 1,
            planting: 0
        };

        // Calculate initial planting
        if (config.plantingMethod === 'percent' && newItem.type === 'plant') {
            newItem.planting = newItem.price * (config.plantingValue / 100);
        }

        estimate.push(newItem);
        saveState();
        renderEstimate();
        showToast('Добавлено в сметную таблицу');
    }

    function renderEstimate() {
        if (estimate.length === 0) {
            emptyState.classList.remove('hidden');
            estimateTable.innerHTML = '';
            generatePdfBtn.disabled = true;
            generateXlsxBtn.disabled = true;
        } else {
            emptyState.classList.add('hidden');
            previewBtn.disabled = false;
            generatePdfBtn.disabled = false;
            generateXlsxBtn.disabled = false;

            estimateTable.innerHTML = estimate.map((item, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td>
                        <select class="item-type-select" data-id="${item.id}">
                            <option value="plant" ${item.type === 'plant' ? 'selected' : ''}>Растение</option>
                            <option value="material" ${item.type === 'material' ? 'selected' : ''}>Материал</option>
                        </select>
                    </td>
                    <td>${item.name}</td>
                    <td>
                        <input type="text" class="item-param-input" style="width:100px; padding:6px; border:1px solid #E2E8F0; border-radius:4px;" value="${item.parameters}" data-id="${item.id}" placeholder="${item.type === 'plant' ? 'Кондиция' : 'Ед. изм.'}">
                    </td>
                    <td>
                        <input type="number" class="item-qty-input" value="${item.quantity}" min="1" data-id="${item.id}">
                    </td>
                    <td>
                        <input type="number" class="item-price-input" value="${item.price.toFixed(2)}" min="0" step="0.01" data-id="${item.id}" style="width:80px; padding:6px; border:1px solid #E2E8F0; border-radius:4px;">
                    </td>
                    <td>${item.planting.toFixed(2)} ₽</td>
                    <td><strong>${((item.price + item.planting) * item.quantity).toFixed(2)} ₽</strong></td>
                    <td>
                        <button class="remove-btn" data-id="${item.id}">
                            <i data-lucide="trash-2"></i>
                        </button>
                    </td>
                </tr>
            `).join('');

            lucide.createIcons();
            attachTableListeners();
        }
        updateTotals();
    }

    function attachTableListeners() {
        document.querySelectorAll('.item-param-input').forEach(input => {
            input.addEventListener('change', (e) => {
                const id = parseInt(e.target.dataset.id);
                const item = estimate.find(i => i.id === id);
                if (item) {
                    item.parameters = e.target.value || '';
                    saveState();
                }
            });
        });

        document.querySelectorAll('.item-qty-input').forEach(input => {
            // input: сохраняем немедленно при каждом нажатии клавиши (без ре-рендера)
            input.addEventListener('input', (e) => {
                const id = parseInt(e.target.dataset.id);
                const item = estimate.find(i => i.id === id);
                if (item) {
                    const newQty = parseInt(e.target.value) || 1;
                    item.quantity = newQty;
                    if (item.type === 'plant' && config.plantingMethod === 'percent') {
                        item.planting = item.price * (config.plantingValue / 100);
                    }
                    saveState();
                    updateTotals();
                }
            });
            // change: полный ре-рендер при потере фокуса (обновляет итоговую колонку в строке)
            input.addEventListener('change', (e) => {
                const id = parseInt(e.target.dataset.id);
                const item = estimate.find(i => i.id === id);
                if (item) {
                    item.quantity = parseInt(e.target.value) || 1;
                    if (item.type === 'plant' && config.plantingMethod === 'percent') {
                        item.planting = item.price * (config.plantingValue / 100);
                    }
                    saveState();
                    renderEstimate();
                }
            });
        });

        document.querySelectorAll('.item-price-input').forEach(input => {
            input.addEventListener('change', (e) => {
                const id = parseInt(e.target.dataset.id);
                const item = estimate.find(i => i.id === id);
                if (item) {
                    item.price = parseFloat(e.target.value) || 0;
                    if (item.type === 'plant' && config.plantingMethod === 'percent') {
                        item.planting = item.price * (config.plantingValue / 100);
                    }
                    saveState();
                    renderEstimate();
                }
            });
        });

        document.querySelectorAll('.item-type-select').forEach(select => {
            select.addEventListener('change', (e) => {
                const id = parseInt(e.target.dataset.id);
                const item = estimate.find(i => i.id === id);
                if (item) {
                    item.type = e.target.value;
                    if (item.type === 'material') {
                        item.planting = 0;
                    } else if (config.plantingMethod === 'percent') {
                        item.planting = item.price * (config.plantingValue / 100);
                    }
                    saveState();
                    renderEstimate();
                }
            });
        });

        document.querySelectorAll('.remove-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const id = parseInt(btn.dataset.id);
                estimate = estimate.filter(i => i.id !== id);
                saveState();
                renderEstimate();
            });
        });
    }

    function updateTotals() {
        let plantsTotal = 0;
        let materialsTotal = 0;
        let laborTotal = 0;

        estimate.forEach(item => {
            if (item.type === 'material') {
                materialsTotal += (item.price * item.quantity);
            } else {
                plantsTotal += (item.price * item.quantity);
                laborTotal += (item.planting * item.quantity);
            }
        });

        const goodsTotal = plantsTotal + materialsTotal;
        const deliveryCost = parseFloat(document.getElementById('delivery-cost')?.value) || 0;
        const subtotal = goodsTotal + laborTotal + deliveryCost;
        const taxAmount = subtotal * (config.taxRate / 100);
        const grandTotal = subtotal + taxAmount;

        const currencyFormat = { minimumFractionDigits: 2, maximumFractionDigits: 2 };
        sumMaterialsEl.innerText = `${goodsTotal.toLocaleString('ru-RU', currencyFormat)} ₽`;
        sumLaborEl.innerText = `${laborTotal.toLocaleString('ru-RU', currencyFormat)} ₽`;
        const sumDeliveryEl = document.getElementById('sum-delivery');
        if (sumDeliveryEl) sumDeliveryEl.innerText = `${deliveryCost.toLocaleString('ru-RU', currencyFormat)} ₽`;
        sumTaxEl.innerText = `${taxAmount.toLocaleString('ru-RU', currencyFormat)} ₽`;
        grandTotalEl.innerText = `${grandTotal.toLocaleString('ru-RU', currencyFormat)} ₽`;
        displayTaxEl.innerText = config.taxRate;
    }

    // Settings Updates
    document.getElementById('tax-rate').addEventListener('input', (e) => {
        config.taxRate = parseFloat(e.target.value) || 0;
        saveState();
        updateTotals();
    });

    const plantingPercentInput = document.getElementById('planting-percent');
    if (plantingPercentInput) {
        plantingPercentInput.addEventListener('input', (e) => {
            config.plantingValue = parseFloat(e.target.value) || 0;
            estimate.forEach(item => {
                if (item.type === 'plant') {
                    item.planting = item.price * (config.plantingValue / 100);
                }
            });
            saveState();
            renderEstimate();
        });
    }

    const deliveryCostInput = document.getElementById('delivery-cost');
    if (deliveryCostInput) {
        deliveryCostInput.addEventListener('input', () => {
            updateTotals();
        });
    }

    // Очистить / Восстановить сметы
    const clearBtn = document.getElementById('clear-estimate');
    const restoreBtn = document.getElementById('restore-estimate');

    clearBtn.addEventListener('click', () => {
        if (estimate.length === 0) return;
        localStorage.setItem('rastenika_estimate_backup', JSON.stringify(estimate));
        estimate = [];
        saveState();
        renderEstimate();
        restoreBtn.classList.remove('hidden');
        showToast('Смета очищена');
    });

    restoreBtn.addEventListener('click', () => {
        try {
            const backup = localStorage.getItem('rastenika_estimate_backup');
            if (backup) {
                estimate = JSON.parse(backup);
                saveState();
                renderEstimate();
                restoreBtn.classList.add('hidden');
                showToast('Смета восстановлена');
            }
        } catch (e) {
            showToast('Не удалось восстановить', 'error');
        }
    });

    // Helper to generate payload
    function buildPayload() {
        const justMaterials = estimate.filter(i => i.type === 'material');
        const justPlants = estimate.filter(i => i.type === 'plant');

        const materialOnlyTotal = justMaterials.reduce((acc, i) => acc + (i.price * i.quantity), 0);

        let allGoodsTotal = 0;
        let laborTotal = 0;

        estimate.forEach(item => {
            if (item.type === 'material') {
                allGoodsTotal += (item.price * item.quantity);
            } else {
                allGoodsTotal += (item.price * item.quantity);
                laborTotal += ((item.planting || 0) * item.quantity);
            }
        });

        const deliveryCost = parseFloat(document.getElementById('delivery-cost')?.value) || 0;
        const subtotal = Math.round((allGoodsTotal + laborTotal + deliveryCost) * 100) / 100;
        const taxAmount = Math.round((subtotal * (config.taxRate / 100)) * 100) / 100;

        return {
            order_number: Date.now(),
            client_name: document.getElementById('company-name') ? document.getElementById('company-name').value || 'Уважаемый клиент' : 'Уважаемый клиент',
            date: new Date().toLocaleDateString('ru-RU'),
            items: justPlants.map(i => ({
                name: i.name,
                parameters: i.parameters,
                quantity: i.quantity,
                price: i.price,
                planting: i.planting || 0,
                total: Math.round((i.price * i.quantity) * 100) / 100
            })),
            materials: justMaterials.map(i => ({
                name: i.name,
                parameters: i.parameters,
                quantity: i.quantity,
                price: i.price,
                total: Math.round((i.price * i.quantity) * 100) / 100
            })),
            material_total: Math.round(materialOnlyTotal * 100) / 100,
            labor_total: Math.round(laborTotal * 100) / 100,
            delivery_total: deliveryCost,
            tax_rate: config.taxRate,
            tax_amount: taxAmount,
            subtotal: subtotal,
            grand_total: Math.round((subtotal + taxAmount) * 100) / 100
        };
    }

    // XLSX Generation
    generateXlsxBtn.addEventListener('click', async () => {
        generateXlsxBtn.disabled = true;
        const originalText = generateXlsxBtn.innerHTML;
        generateXlsxBtn.innerHTML = '<i data-lucide="loader" class="spin"></i> Загрузка...';
        lucide.createIcons();

        const payload = buildPayload();

        showToast('Выгрузка XLSX...');

        try {
            const response = await fetch('/generate-xlsx', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = `Смета_Ландшафт_${new Date().getTime()}.xlsx`;
                document.body.appendChild(a);
                a.click();
                setTimeout(() => {
                    window.URL.revokeObjectURL(url);
                    a.remove();
                }, 100);
                showToast('XLSX успешно выгружен!');
            } else {
                const errorData = await response.json();
                showToast(errorData.error || 'Ошибка при выгрузке XLSX', 'error');
            }
        } catch (err) {
            showToast('Сетевая ошибка', 'error');
        } finally {
            generateXlsxBtn.disabled = false;
            generateXlsxBtn.innerHTML = originalText;
            lucide.createIcons();
        }
    });

    // PDF Generation
    generatePdfBtn.addEventListener('click', async () => {
        generatePdfBtn.disabled = true;
        const originalText = generatePdfBtn.innerHTML;
        generatePdfBtn.innerHTML = '<i data-lucide="loader" class="spin"></i> Генерация...';
        lucide.createIcons();

        const payload = buildPayload();

        showToast('Генерация PDF...');

        try {
            const response = await fetch('/generate-pdf', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = `Смета_Ландшафт_${new Date().getTime()}.pdf`;
                document.body.appendChild(a);
                a.click();
                setTimeout(() => {
                    window.URL.revokeObjectURL(url);
                    a.remove();
                }, 100);
                showToast('PDF успешно сформирован!');
            } else {
                const errorData = await response.json();
                showToast(errorData.error || 'Ошибка при генерации PDF', 'error');
            }
        } catch (err) {
            showToast('Сетевая ошибка', 'error');
        } finally {
            generatePdfBtn.disabled = false;
            generatePdfBtn.innerHTML = originalText;
            lucide.createIcons();
        }
    });

    // Preview Logic
    previewBtn.addEventListener('click', async () => {
        previewBtn.disabled = true;
        const originalText = previewBtn.innerHTML;
        previewBtn.innerHTML = '<i data-lucide="loader" class="spin"></i> Загрузка...';
        lucide.createIcons();

        try {
            const payload = buildPayload();
            const response = await fetch('/generate-pdf', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                previewIframe.src = url;
                previewModal.classList.remove('hidden');
            } else {
                const errorData = await response.json();
                showToast(errorData.error || 'Ошибка при генерации превью', 'error');
            }
        } catch (err) {
            showToast('Сетевая ошибка при загрузке превью', 'error');
        } finally {
            previewBtn.disabled = false;
            previewBtn.innerHTML = originalText;
            lucide.createIcons();
        }
    });

    closeModalBtn.addEventListener('click', () => {
        previewModal.classList.add('hidden');
        previewIframe.src = '';
    });

    // Manual Add Item Logic
    const manualAddBtn = document.getElementById('manual-add-btn');
    const manualAddModal = document.getElementById('manual-add-modal');
    const manualAddForm = document.getElementById('manual-add-form');
    const closeManualModalBtn = document.getElementById('close-manual-modal');
    const cancelManualBtn = document.getElementById('cancel-manual-btn');

    manualAddBtn.addEventListener('click', () => {
        manualAddModal.classList.remove('hidden');
        document.getElementById('manual-name').focus();
    });

    closeManualModalBtn.addEventListener('click', () => {
        manualAddModal.classList.add('hidden');
        manualAddForm.reset();
    });

    cancelManualBtn.addEventListener('click', () => {
        manualAddModal.classList.add('hidden');
        manualAddForm.reset();
    });

    // Close modal on overlay click
    manualAddModal.addEventListener('click', (e) => {
        if (e.target === manualAddModal) {
            manualAddModal.classList.add('hidden');
            manualAddForm.reset();
        }
    });

    manualAddForm.addEventListener('submit', (e) => {
        e.preventDefault();

        const name = document.getElementById('manual-name').value.trim();
        const params = document.getElementById('manual-params').value.trim();
        const type = document.getElementById('manual-type').value;
        const price = parseFloat(document.getElementById('manual-price').value) || 0;
        const qty = parseFloat(document.getElementById('manual-qty').value) || 1;

        if (!name) {
            showToast('Укажите наименование', 'error');
            return;
        }

        const newItem = {
            id: Date.now(),
            type: type,
            name: name,
            parameters: params,
            price: price,
            quantity: qty,
            planting: 0
        };

        // Calculate planting for plants
        if (type === 'plant' && config.plantingMethod === 'percent') {
            newItem.planting = newItem.price * (config.plantingValue / 100);
        }

        estimate.push(newItem);
        saveState();
        renderEstimate();
        manualAddModal.classList.add('hidden');
        manualAddForm.reset();
        showToast('Позиция добавлена в смету');
    });

    // Toast Notification
    function showToast(message, type = 'success') {
        const toast = document.getElementById('toast');
        toast.innerText = message;
        toast.className = `toast ${type}`;
        toast.classList.remove('hidden');
        setTimeout(() => toast.classList.add('hidden'), 3000);
    }
});
