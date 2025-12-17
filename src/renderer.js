import * as XLSX from 'xlsx-js-style';

// 常量定义
const MODEL_ORDER = ['系统', '砖机', '摆渡车', '辅机', '运输车']; // 机型排序顺序
const CONCURRENCY_LIMIT = 10; // 并发处理文件数量限制
const API_BASE_URL = '/api'; // API 基础 URL

// 全局变量
let devices = new Map(); // 设备清单，key: 物料编号
let selectedFiles = []; // 选择的文件
let unknownDevices = []; // 未识别的设备
let vulnerableList = []; // 易损清单结果
let matchedDevices = []; // 白名单已匹配设备
let allFiles = []; // 处理的所有文件

// DOM元素缓存
const dom = {
    // 标签页相关
    tabs: document.querySelectorAll('.tab'),
    tabContents: document.querySelectorAll('.tab-content'),
    
    // 设备管理相关
    materialIdInput: document.getElementById('materialId'),
    descriptionInput: document.getElementById('description'),
    spareCountInput: document.getElementById('spareCount'),
    unitInput: document.getElementById('unit'),
    remarkInput: document.getElementById('remark'),
    statusSelect: document.getElementById('status'),
    deviceTypeFilter: document.getElementById('deviceTypeFilter'),
    materialIdSearch: document.getElementById('materialIdSearch'),
    addDeviceBtn: document.getElementById('addDevice'),
    saveDevicesBtn: document.getElementById('saveDevices'),
    loadDevicesBtn: document.getElementById('loadDevices'),
    deviceTableBody: document.getElementById('deviceTableBody'),
    
    // 文件处理相关
    selectFolderBtn: document.getElementById('selectFolder'),
    processFilesBtn: document.getElementById('processFiles'),
    exportListBtn: document.getElementById('exportList'),
    fileInfoDiv: document.getElementById('fileInfo'),
    summaryDiv: document.getElementById('summary'),
    
    // 结果展示相关
    matchedSection: document.getElementById('matchedSection'),
    matchedTableBody: document.getElementById('matchedTableBody'),
    unknownSection: document.getElementById('unknownSection'),
    unknownTableBody: document.getElementById('unknownTableBody'),
    resultSection: document.getElementById('resultSection'),
    resultTableBody: document.getElementById('resultTableBody')
};

// 初始化：从服务器加载设备清单并添加事件监听器
function initApp() {
    // 加载设备清单
    loadDevices();
    
    // 添加事件监听器
    addEventListeners();
}

// 添加所有事件监听器
function addEventListeners() {
    // 标签页切换事件
    dom.tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            switchTab(tab.dataset.tab);
        });
    });
    
    // 设备管理事件
    dom.addDeviceBtn.addEventListener('click', addDevice);
    dom.saveDevicesBtn.addEventListener('click', saveDevices);
    dom.loadDevicesBtn.addEventListener('click', loadDevices);
    dom.deviceTypeFilter.addEventListener('change', renderDeviceTable);
    dom.materialIdSearch.addEventListener('input', renderDeviceTable);
    
    // 文件处理事件
    dom.selectFolderBtn.addEventListener('click', selectFiles);
    dom.processFilesBtn.addEventListener('click', processFiles);
    dom.exportListBtn.addEventListener('click', exportList);
    
    // 批量导入/导出事件
    document.getElementById('importBlacklist').addEventListener('click', importBlacklist);
    document.getElementById('importWhitelist').addEventListener('click', importWhitelist);
    document.getElementById('exportBlacklist').addEventListener('click', exportBlacklist);
    document.getElementById('exportVulnerableLibrary').addEventListener('click', exportVulnerableLibrary);
}

// 标签页切换
function switchTab(tabName) {
    dom.tabs.forEach(tab => tab.classList.remove('active'));
    dom.tabContents.forEach(content => content.classList.remove('active'));
    
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
    document.getElementById(tabName).classList.add('active');
}



// 设备清单维护
async function addDevice() {
    const materialId = dom.materialIdInput.value.trim();
    const description = dom.descriptionInput.value.trim();
    const spareCount = parseInt(dom.spareCountInput.value) || 0;
    const unit = dom.unitInput.value.trim();
    const remark = dom.remarkInput.value.trim();
    const status = dom.statusSelect.value;
    
    if (!materialId) {
        alert('物料编号不能为空');
        return;
    }
    
    // 直接使用materialId作为唯一键，不再使用model
    devices.set(materialId, {
        materialId,
        description,
        spareCount,
        unit,
        remark,
        status
    });
    
    // 清空输入框
    dom.materialIdInput.value = '';
    dom.descriptionInput.value = '';
    dom.spareCountInput.value = '';
    dom.unitInput.value = '';
    dom.remarkInput.value = '';
    
    // 重新渲染设备表格
    renderDeviceTable();
    
    // 保存到数据库
    try {
        const devicesArray = Array.from(devices.values());
        const response = await fetch(`${API_BASE_URL}/devices`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(devicesArray)
        });
        
        const result = await response.json();
        if (response.ok && result.success) {
            alert('设备添加成功');
        } else {
            alert('设备保存失败: ' + (result.error || '未知错误'));
        }
    } catch (error) {
        console.error('保存设备清单失败:', error);
        alert('设备保存失败: ' + error.message);
    }
}

function renderDeviceTable() {
    // 获取当前过滤条件
    const filterType = dom.deviceTypeFilter.value;
    const searchTerm = dom.materialIdSearch.value.trim().toLowerCase();
    
    // 过滤设备
    const filteredDevices = [];
    for (const [materialId, device] of devices.entries()) {
        const matchesType = filterType === 'all' || device.status === filterType;
        const matchesSearch = !searchTerm || device.materialId.toLowerCase().includes(searchTerm);
        
        if (matchesType && matchesSearch) {
            filteredDevices.push(device);
        }
    }
    
    // 按状态分类设备
    const whiteListDevices = filteredDevices.filter(device => device.status === '白名单');
    const blackListDevices = filteredDevices.filter(device => device.status === '黑名单');
    
    // 获取设备表的表头
    const deviceTable = document.getElementById('deviceTable');
    const thead = deviceTable.querySelector('thead');
    
    // 动态生成表头
    if (filterType === '黑名单' && blackListDevices.length > 0) {
        // 黑名单只显示4列表头
        thead.innerHTML = `
            <tr>
                <th>物料编号</th>
                <th>物料描述</th>
                <th>状态</th>
                <th>操作</th>
            </tr>
        `;
    } else {
        // 白名单和全部设备显示完整7列表头（移除了机型列）
        thead.innerHTML = `
            <tr>
                <th>物料编号</th>
                <th>物料描述</th>
                <th>备件数</th>
                <th>单位</th>
                <th>备注</th>
                <th>状态</th>
                <th>操作</th>
            </tr>
        `;
    }
    
    // 清空表格内容
    dom.deviceTableBody.innerHTML = '';
    
    // 渲染白名单设备
    if (whiteListDevices.length > 0) {
        whiteListDevices.forEach(device => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${device.materialId}</td>
                <td>${device.description || ''}</td>
                <td>${device.spareCount}</td>
                <td>${device.unit}</td>
                <td>${device.remark}</td>
                <td>
                    <span class="status-badge status-white">
                        ${device.status}
                    </span>
                </td>
                <td>
                    <button class="danger" style="padding: 5px 10px; font-size: 12px; margin-right: 5px;">
                        删除
                    </button>
                    <button class="secondary" style="padding: 5px 10px; font-size: 12px;">
                        改为黑名单
                    </button>
                </td>
            `;
            dom.deviceTableBody.appendChild(row);
            
            // 直接使用materialId作为唯一键
            const buttons = row.querySelectorAll('button');
            buttons[0].addEventListener('click', () => removeDevice(device.materialId));
            buttons[1].addEventListener('click', () => toggleDeviceStatus(device.materialId));
        });
    }
    
    // 渲染黑名单设备 - 只显示4列
    if (blackListDevices.length > 0) {
        blackListDevices.forEach(device => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${device.materialId}</td>
                <td>${device.description || ''}</td>
                <td>
                    <span class="status-badge status-black">
                        ${device.status}
                    </span>
                </td>
                <td>
                    <button class="danger" style="padding: 5px 10px; font-size: 12px; margin-right: 5px;">
                        删除
                    </button>
                    <button class="secondary" style="padding: 5px 10px; font-size: 12px;">
                        改为白名单
                    </button>
                </td>
            `;
            dom.deviceTableBody.appendChild(row);
            
            // 直接使用materialId作为唯一键
            const buttons = row.querySelectorAll('button');
            buttons[0].addEventListener('click', () => removeDevice(device.materialId));
            buttons[1].addEventListener('click', () => toggleDeviceStatus(device.materialId));
        });
    }
    
    // 空状态处理
    if (filteredDevices.length === 0) {
        const emptyRow = document.createElement('tr');
        let emptyMessage = '暂无设备数据';
        if (filterType === '白名单') {
            emptyMessage = '暂无白名单设备数据';
        } else if (filterType === '黑名单') {
            emptyMessage = '暂无黑名单设备数据';
        }
        const colSpan = filterType === '黑名单' ? 4 : 7;
        emptyRow.innerHTML = `
            <td colspan="${colSpan}" style="text-align: center; padding: 20px; color: #999;">
                ${emptyMessage}
            </td>
        `;
        dom.deviceTableBody.appendChild(emptyRow);
    }
}

async function removeDevice(materialId) {
    try {
        // 从内存中删除设备
        devices.delete(materialId);
        
        // 重新渲染设备表格
        renderDeviceTable();
        
        // 保存更新后的设备清单到数据库
        const devicesArray = Array.from(devices.values());
        const response = await fetch(`${API_BASE_URL}/devices`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(devicesArray)
        });
        
        const result = await response.json();
        if (response.ok && result.success) {
            alert('设备删除成功');
        } else {
            console.error('保存设备清单失败:', result.error || '未知错误');
            alert('删除设备失败: ' + (result.error || '未知错误'));
        }
    } catch (error) {
        console.error('删除设备失败:', error);
        alert('删除设备失败: ' + error.message);
    }
}

// 暴露到全局作用域，供HTML onclick事件使用
window.removeDevice = removeDevice;

// 切换设备状态（白名单 ↔ 黑名单）
async function toggleDeviceStatus(materialId) {
    try {
        const device = devices.get(materialId);
        if (!device) {
            alert('未找到该设备');
            return;
        }
        
        // 切换状态
        const newStatus = device.status === '白名单' ? '黑名单' : '白名单';
        
        // 更新设备对象
        const updatedDevice = {
            ...device,
            status: newStatus
        };
        
        // 更新内存中的设备清单
        devices.set(materialId, updatedDevice);
        
        // 重新渲染设备表格
        renderDeviceTable();
        
        // 保存到数据库
        const devicesArray = Array.from(devices.values());
        const response = await fetch(`${API_BASE_URL}/devices`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(devicesArray)
        });
        
        const result = await response.json();
        if (response.ok && result.success) {
            alert(`设备已从${device.status}改为${newStatus}`);
        } else {
            console.error('保存设备清单失败:', result.error || '未知错误');
            alert(`设备状态已更改，但保存失败: ${result.error || '未知错误'}`);
        }
    } catch (error) {
        console.error('切换设备状态失败:', error);
        alert('切换设备状态失败: ' + error.message);
    }
}

// 暴露到全局作用域，供HTML onclick事件使用
window.toggleDeviceStatus = toggleDeviceStatus;

async function saveDevices() {
    try {
        const devicesArray = Array.from(devices.values());
        const response = await fetch(`${API_BASE_URL}/devices`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(devicesArray)
        });
        
        const result = await response.json();
        if (response.ok && result.success) {
            alert('设备清单保存成功');
        } else {
            alert('设备清单保存失败: ' + (result.error || '未知错误'));
        }
    } catch (error) {
        console.error('保存设备清单失败:', error);
        alert('保存设备清单失败: ' + error.message);
    }
}

async function loadDevices() {
    try {
        const response = await fetch(`${API_BASE_URL}/devices`);
        if (!response.ok) {
            throw new Error('网络请求失败');
        }
        
        const devicesArray = await response.json();
        
        // 将设备数组转换为 Map 对象，直接使用materialId作为唯一键
        // 后端已经确保每个materialId只返回一条记录，所以这里可以直接设置
        const devicesMap = new Map();
        devicesArray.forEach(device => {
            // 直接设置，后端已经确保每个materialId唯一
            devicesMap.set(device.materialId, device);
        });
        
        devices = devicesMap;
        renderDeviceTable();
        alert('设备清单加载成功');
    } catch (error) {
        console.error('加载设备清单失败:', error);
        alert('加载设备清单失败: ' + error.message);
    }
}

// 选择文件
function selectFiles() {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = '.xlsx,.xls';
    input.onchange = (e) => {
            selectedFiles = Array.from(e.target.files);
            if (selectedFiles.length > 0) {
                dom.fileInfoDiv.innerHTML = `<strong>已选择文件:</strong> ${selectedFiles.length} 个<br>
                                         ${selectedFiles.map(file => `• ${file.name}`).join('<br>')}`;
                dom.fileInfoDiv.style.display = 'block';
            }
        };
    input.click();
}

// 处理Excel文件
async function processFiles() {
    if (selectedFiles.length === 0) {
        alert('请先选择Excel文件');
        return;
    }
    
    // 确保设备清单已加载
    if (devices.size === 0) {
        const response = await fetch(`${API_BASE_URL}/devices`);
        if (!response.ok) {
            throw new Error('网络请求失败');
        }
        
        const devicesArray = await response.json();
        
        // 将设备数组转换为 Map 对象，直接使用materialId作为唯一键
        const devicesMap = new Map();
        devicesArray.forEach(device => {
            devicesMap.set(device.materialId, device);
        });
        
        devices = devicesMap;
        if (devices.size === 0) {
            alert('请先维护设备清单');
            return;
        }
    }
    
    // 重置结果数组
    unknownDevices = [];
    vulnerableList = [];
    matchedDevices = [];
    
    // 不再在处理文件时进行去重，只在最后结果生成时进行一次去重
    
    let totalFiles = selectedFiles.length;
    let processedFiles = 0;
    let totalDevices = 0;
    let matchedWhite = 0;
    let matchedBlack = 0;
    let unmatched = 0;
    
    // 预编译ERP字段名称匹配正则
    const erpFieldPatterns = [
        /ERP/i,
        /ERP编号/i,
        /物料号/i,
        /ERP编码/i,
        /MaterialID/i,
        /material_id/i
    ];
    
    // 并发控制：使用常量控制同时处理的文件数量
    const processResults = [];
    
    // 并发处理函数
    async function processFilesConcurrently() {
        for (let i = 0; i < selectedFiles.length; i += CONCURRENCY_LIMIT) {
            const batch = selectedFiles.slice(i, i + CONCURRENCY_LIMIT);
            const batchPromises = batch.map(async (file) => {
                try {
                    // 读取Excel文件
                    const arrayBuffer = await file.arrayBuffer();
                    const workbook = XLSX.read(arrayBuffer, { cellDates: true, cellText: false });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    
                    // 优化：读取需要的列，减少内存占用
                    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
                    let erpColumnIndex = -1;
                    let descriptionColumnIndex = -1;
                    let unitColumnIndex = -1;
                    
                    // 查找需要的列
                    for (let c = range.s.c; c <= range.e.c; c++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 0, c });
                        const cell = worksheet[cellAddress];
                        if (cell && cell.v) {
                            const header = cell.v.toString();
                            if (erpFieldPatterns.some(pattern => pattern.test(header))) {
                                erpColumnIndex = c;
                            } else if (/物料描述/i.test(header)) {
                                descriptionColumnIndex = c;
                            } else if (/单位/i.test(header)) {
                                unitColumnIndex = c;
                            }
                        }
                    }
                    
                    if (erpColumnIndex === -1) {
                        console.log(`文件 ${file.name} 中未找到ERP列`);
                        return;
                    }
                    
                    // 处理所有行，读取需要的字段
                    const deviceData = [];
                    for (let r = 1; r <= range.e.r; r++) {
                        const erpCell = XLSX.utils.encode_cell({ r, c: erpColumnIndex });
                        const erpCode = worksheet[erpCell]?.v?.toString().trim();
                        
                        if (erpCode) {
                            // 读取物料描述
                            let description = '';
                            if (descriptionColumnIndex !== -1) {
                                const descCell = XLSX.utils.encode_cell({ r, c: descriptionColumnIndex });
                                description = worksheet[descCell]?.v?.toString().trim() || '';
                            }
                            
                            // 跳过物料描述为空的行（设备大类）
                            if (!description) {
                                continue;
                            }
                            
                            // 读取单位
                            let unit = '';
                            if (unitColumnIndex !== -1) {
                                const unitCell = XLSX.utils.encode_cell({ r, c: unitColumnIndex });
                                unit = worksheet[unitCell]?.v?.toString().trim() || '';
                            }
                            
                            deviceData.push({
                                erpCode: erpCode,
                                description: description,
                                unit: unit
                            });
                        }
                    }
                    
                    // 根据文件名判断机型
                    let model = '';
                    const fileName = file.name.toLowerCase();
                    if (fileName.includes('ap')) {
                        model = '系统';
                    } else if (fileName.includes('摆渡车')) {
                        model = '摆渡车';
                    } else if (fileName.includes('运输车')) {
                        model = '运输车';
                    } else if (fileName.includes('砖机')) {
                        model = '砖机';
                    } else if (fileName.includes('辅机')) {
                        model = '辅机';
                    }
                    
                    totalDevices += deviceData.length;
                    
                    // 批量处理ERP编号
                    const fileResults = {
                        unknown: [],
                        vulnerable: [],
                        matchedWhite: 0,
                        matchedBlack: 0,
                        unmatched: 0
                    };
                    
                    deviceData.forEach(deviceItem => {
                        const materialId = deviceItem.erpCode;
                        // 直接使用materialId作为唯一键查找设备
                        const device = devices.get(materialId);
                        
                        if (device) {
                            // 设备已存在 - 先检查黑名单，再检查白名单，确保所有设备都能被处理
                            if (device.status === '黑名单') {
                                // 黑名单优先级更高，直接标记为匹配黑名单
                                fileResults.matchedBlack++;
                            } else if (device.status === '白名单') {
                                // 白名单设备，直接添加到易损清单
                                fileResults.vulnerable.push({
                                    erpCode: materialId,
                                    ...device,
                                    model: model // 使用根据文件名判断的机型，而不是设备清单中的机型
                                });
                                fileResults.matchedWhite++;
                            } else {
                                // 状态异常的设备，添加到未知设备列表
                                console.log(`发现状态异常设备: ${materialId}, 状态: ${device.status}`);
                                fileResults.unknown.push({
                                    erpCode: materialId,
                                    materialId: device.materialId,
                                    description: device.description || deviceItem.description,
                                    spareCount: device.spareCount || 0,
                                    unit: device.unit || deviceItem.unit,
                                    model: model, // 根据文件名判断的机型
                                    remark: device.remark || '',
                                    isVulnerable: true
                                });
                                fileResults.unmatched++;
                            }
                        } else {
                            // 设备不存在，添加到未知设备
                            fileResults.unknown.push({
                                erpCode: materialId,
                                materialId: '',
                                description: deviceItem.description,
                                spareCount: 0,
                                unit: deviceItem.unit,
                                model: model, // 根据文件名判断的机型
                                remark: '',
                                isVulnerable: true
                            });
                            fileResults.unmatched++;
                        }
                    });
                    
                    return fileResults;
                } catch (error) {
                    console.error(`处理文件 ${file.name} 失败:`, error);
                    return null;
                } finally {
                    processedFiles++;
                }
            });
            
            // 等待当前批次处理完成
            const batchResults = await Promise.all(batchPromises);
            processResults.push(...batchResults);
        }
    }
    
    // 执行并发处理
    await processFilesConcurrently();
    
    const results = processResults;
    
    // 合并结果
    let totalVulnerableBeforeMerge = 0;
    let totalUnknownBeforeMerge = 0;
    results.forEach(result => {
        if (result) {
            totalVulnerableBeforeMerge += result.vulnerable.length;
            totalUnknownBeforeMerge += result.unknown.length;
            vulnerableList.push(...result.vulnerable);
            unknownDevices.push(...result.unknown);
            matchedWhite += result.matchedWhite;
            matchedBlack += result.matchedBlack;
            unmatched += result.unmatched;
        }
    });
    
    console.log(`合并结果前：`);
    console.log(`  单个文件易损设备总数：${totalVulnerableBeforeMerge}`);
    console.log(`  单个文件未知设备总数：${totalUnknownBeforeMerge}`);
    console.log(`  合并后易损设备总数：${vulnerableList.length}`);
    console.log(`  合并后未知设备总数：${unknownDevices.length}`);
    console.log(`  匹配白名单数量：${matchedWhite}`);
    console.log(`  匹配黑名单数量：${matchedBlack}`);
    console.log(`  未匹配数量：${unmatched}`);
    
    // 去重处理 - 未知设备按机型和ERP编号去重，易损清单只按ERP编号去重
    const unknownBeforeDedup = unknownDevices.length;
    unknownDevices = uniqueByKey(unknownDevices, ['model', 'erpCode']);
    console.log(`未知设备去重：${unknownBeforeDedup} → ${unknownDevices.length}（去重${unknownBeforeDedup - unknownDevices.length}个）`);
    
    // 直接使用Set来实现去重，只保留唯一的ERP编号
    // 这里必须确保去重逻辑正确执行，解决跨文件重复的问题
    const vulnerableBeforeDedup = vulnerableList.length;
    const seenErpCodes = new Set();
    vulnerableList = vulnerableList.filter(device => {
        // 使用erpCode作为唯一标识进行去重
        const erpCode = device.erpCode;
        if (seenErpCodes.has(erpCode)) {
            console.log(`易损清单去重：重复的物料编号 ${erpCode}，已去重`);
            return false;
        }
        seenErpCodes.add(erpCode);
        return true;
    });
    console.log(`易损清单去重：${vulnerableBeforeDedup} → ${vulnerableList.length}（去重${vulnerableBeforeDedup - vulnerableList.length}个）`);
    
    // 统计各机型设备数量
    const devicesByModel = new Map();
    vulnerableList.forEach(device => {
        const model = device.model;
        devicesByModel.set(model, (devicesByModel.get(model) || 0) + 1);
    });
    console.log('易损清单各机型数量：');
    devicesByModel.forEach((count, model) => {
        console.log(`  ${model}：${count}个`);
    });
    
    // 再次验证去重结果，确保没有重复
    const finalSeenErpCodes = new Set();
    let hasDuplicates = false;
    vulnerableList.forEach(device => {
        if (finalSeenErpCodes.has(device.erpCode)) {
            console.error(`严重错误：去重后仍有重复的物料编号: ${device.erpCode}`);
            hasDuplicates = true;
        }
        finalSeenErpCodes.add(device.erpCode);
    });
    
    if (!hasDuplicates) {
        console.log('去重成功，易损清单中没有重复的物料编号');
    } else {
        // 如果仍有重复，强制再次去重
        console.log('发现重复，正在进行强制去重');
        const finalUniqueSet = new Map();
        vulnerableList.forEach(device => {
            finalUniqueSet.set(device.erpCode, device);
        });
        vulnerableList = Array.from(finalUniqueSet.values());
        console.log('强制去重完成，当前易损清单数量:', vulnerableList.length);
    }
    
    // 处理白名单匹配设备 - 从易损清单中提取，只显示4列
    matchedDevices = vulnerableList.map(device => ({
        erpCode: device.erpCode,
        description: device.description,
        model: device.model
    }));
    
    // 更新统计信息
    dom.summaryDiv.innerHTML = `
        <strong>处理结果:</strong><br>
        - 总文件数: ${totalFiles} 个<br>
        - 总设备数: ${totalDevices} 个<br>
        - 匹配白名单: ${matchedWhite} 个<br>
        - 匹配黑名单: ${matchedBlack} 个<br>
        - 已匹配白名单设备(去重后): ${matchedDevices.length} 个<br>
        - 未识别设备: ${unknownDevices.length} 个 (去重后)<br>
    `;
    dom.summaryDiv.style.display = 'block';
    
    // 显示白名单匹配设备
    if (matchedDevices.length > 0) {
        dom.matchedSection.style.display = 'block';
        renderMatchedTable();
    } else {
        dom.matchedSection.style.display = 'none';
    }
    
    // 显示未识别设备
    if (unknownDevices.length > 0) {
        dom.unknownSection.style.display = 'block';
        renderUnknownTable();
    } else {
        dom.unknownSection.style.display = 'none';
        renderResultTable();
        dom.exportListBtn.disabled = false;
    }
    
    dom.resultSection.style.display = 'block';
}

function renderUnknownTable() {
    dom.unknownTableBody.innerHTML = '';
    
    unknownDevices.forEach((device, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${device.erpCode}</td>
            <td><input type="text" id="materialId_${index}" value="${device.description || device.erpCode}" placeholder="物料描述" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 300px; min-width: 200px;"></td>
            <td><input type="number" id="spareCount_${index}" value="${device.spareCount}" min="0" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 80px;"></td>
            <td><input type="text" id="unit_${index}" value="${device.unit}" placeholder="单位" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 80px;"></td>
            <td><input type="text" id="remark_${index}" value="${device.remark}" placeholder="备注" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 120px;"></td>
            <td>
                <select id="isVulnerable_${index}" onchange="toggleDeviceInfo(${index}, this.value)" style="width: 80px;">
                    <option value="true" ${device.isVulnerable ? 'selected' : ''}>是</option>
                    <option value="false" ${!device.isVulnerable ? 'selected' : ''}>否</option>
                </select>
            </td>
        `;
        dom.unknownTableBody.appendChild(row);
    });
    
    // 添加确认按钮
    const confirmRow = document.createElement('tr');
    confirmRow.innerHTML = `
        <td colspan="6" style="text-align: center;">
            <button onclick="confirmUnknownDevices()" style="margin-top: 10px;">
                确认选择
            </button>
        </td>
    `;
    dom.unknownTableBody.appendChild(confirmRow);
}

// 切换设备信息输入框显示
function toggleDeviceInfo(index, isVulnerable) {
    const show = isVulnerable === 'true';
    document.getElementById(`materialId_${index}`).style.display = show ? 'block' : 'none';
    document.getElementById(`spareCount_${index}`).style.display = show ? 'block' : 'none';
    document.getElementById(`unit_${index}`).style.display = show ? 'block' : 'none';
    document.getElementById(`remark_${index}`).style.display = show ? 'block' : 'none';
}

// 暴露到全局作用域，供HTML onclick事件使用
window.toggleDeviceInfo = toggleDeviceInfo;

function confirmUnknownDevices() {
    // 更新未知设备信息并处理
    for (let i = 0; i < unknownDevices.length; i++) {
        const device = unknownDevices[i];
        const erpCode = device.erpCode;
        const description = document.getElementById(`materialId_${i}`).value.trim() || device.description || device.erpCode;
        const spareCount = parseInt(document.getElementById(`spareCount_${i}`).value) || 0;
        const unit = document.getElementById(`unit_${i}`).value.trim() || device.unit;
        const remark = document.getElementById(`remark_${i}`).value.trim();
        const isVulnerable = document.getElementById(`isVulnerable_${i}`).value === 'true';
        
        // 使用ERP编号作为materialId
        const materialId = erpCode;
        
        // 验证逻辑
        if (isVulnerable) {
            // 选择是的情况：物料描述、备件数、单位都不能为空，备件数不能为0
            if (!description) {
                alert(`第${i+1}行：物料描述不能为空`);
                return;
            }
            if (spareCount <= 0) {
                alert(`第${i+1}行：备件数必须大于0`);
                return;
            }
            if (!unit) {
                alert(`第${i+1}行：单位不能为空`);
                return;
            }
        }
        
        // 根据是否易损创建设备对象
        const deviceObj = {
            materialId,
            description,
            status: isVulnerable ? '白名单' : '黑名单' // 确保status值有效
        };
        
        // 如果是易损，需要完整字段
        if (isVulnerable) {
            deviceObj.spareCount = spareCount;
            deviceObj.unit = unit;
            deviceObj.remark = remark;
        } else {
            // 选择否的情况：只需要编号和描述
            deviceObj.spareCount = 0;
            deviceObj.unit = '';
            deviceObj.remark = '';
        }
        
        // 直接使用materialId作为唯一键添加到设备清单
        devices.set(materialId, deviceObj);
        
        // 如果是易损，添加到结果
        if (isVulnerable) {
            // 检查该物料编号是否已经存在于易损清单中
            const existingIndex = vulnerableList.findIndex(item => item.erpCode === device.erpCode);
            if (existingIndex === -1) {
                vulnerableList.push({
                    erpCode: device.erpCode,
                    materialId,
                    description,
                    spareCount,
                    unit,
                    remark,
                    status: '白名单'
                });
            }
        }
    }
    
    // 保存更新后的设备清单
    saveDevices();
    
    // 隐藏未知设备区域
    dom.unknownSection.style.display = 'none';
    
    // 使用Map实现严格去重，确保易损清单中物料编号唯一
    const uniqueDevicesMap = new Map();
    vulnerableList.forEach(device => {
        uniqueDevicesMap.set(device.erpCode, device);
    });
    vulnerableList = Array.from(uniqueDevicesMap.values());
    
    console.log('未知设备处理后，易损清单去重完成，当前数量:', vulnerableList.length);
    
    // 显示结果
    renderResultTable();
    dom.exportListBtn.disabled = false;
}

// 暴露到全局作用域，供HTML onclick事件使用
window.confirmUnknownDevices = confirmUnknownDevices;

function renderMatchedTable() {
    dom.matchedTableBody.innerHTML = '';
    
    // 更新标题，显示总数
    const h3 = dom.matchedSection.querySelector('h3');
    if (h3) {
        h3.textContent = `白名单已匹配设备 (共 ${matchedDevices.length} 个)`;
    }
    
    // 按机型排序匹配设备
    const sortedMatchedDevices = [...matchedDevices].sort((a, b) => {
        const aIndex = MODEL_ORDER.indexOf(a.model);
        const bIndex = MODEL_ORDER.indexOf(b.model);
        return aIndex - bIndex;
    });
    
    sortedMatchedDevices.forEach(device => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${device.erpCode}</td>
            <td>${device.description}</td>
            <td>${device.model}</td>
        `;
        dom.matchedTableBody.appendChild(row);
    });
}

function renderResultTable() {
    dom.resultTableBody.innerHTML = '';
    
    // 按机型排序易损清单
    const sortedVulnerableList = [...vulnerableList].sort((a, b) => {
        const aIndex = MODEL_ORDER.indexOf(a.model);
        const bIndex = MODEL_ORDER.indexOf(b.model);
        return aIndex - bIndex;
    });
    
    sortedVulnerableList.forEach(device => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${device.erpCode}</td>
            <td>${device.spareCount}</td>
            <td>${device.unit}</td>
            <td>${device.model}</td>
            <td>${device.remark}</td>
        `;
        dom.resultTableBody.appendChild(row);
    });
}

// 导出易损清单
async function exportList() {
    if (vulnerableList.length === 0) {
        alert('没有易损设备可以导出');
        return;
    }
    
    try {
        console.log('开始导出易损清单，原始数据数量:', vulnerableList.length);
        
        // 使用统一的机型排序顺序常量
        
        // 第一步：使用Map进行严格去重，确保每个物料编号只出现一次
        const uniqueDevicesMap = new Map();
        vulnerableList.forEach((device, index) => {
            const materialId = device.erpCode;
            if (uniqueDevicesMap.has(materialId)) {
                console.log(`导出去重：发现重复物料编号 ${materialId}，已忽略重复项（第${index+1}条）`);
                return;
            }
            uniqueDevicesMap.set(materialId, device);
        });
        
        let uniqueVulnerableList = Array.from(uniqueDevicesMap.values());
        console.log('第一步去重后数据数量:', uniqueVulnerableList.length);
        
        // 第二步：再次验证去重结果，确保没有重复
        const finalSeenIds = new Set();
        let hasDuplicates = false;
        uniqueVulnerableList.forEach((device, index) => {
            if (finalSeenIds.has(device.erpCode)) {
                console.error(`严重错误：去重后仍有重复物料编号 ${device.erpCode}（第${index+1}条）`);
                hasDuplicates = true;
            }
            finalSeenIds.add(device.erpCode);
        });
        
        if (hasDuplicates) {
            console.error('去重失败，尝试第三次强制去重');
            // 第三次强制去重
            const finalMap = new Map();
            uniqueVulnerableList.forEach(device => {
                finalMap.set(device.erpCode, device);
            });
            uniqueVulnerableList = Array.from(finalMap.values());
            console.log('第三次去重后数据数量:', uniqueVulnerableList.length);
        }
        
        console.log('导出前去重完成，最终数据数量:', uniqueVulnerableList.length);
        
        // 按机型排序易损清单
        const sortedVulnerableList = [...uniqueVulnerableList].sort((a, b) => {
            const aIndex = MODEL_ORDER.indexOf(a.model);
            const bIndex = MODEL_ORDER.indexOf(b.model);
            return aIndex - bIndex;
        });
        
        // 准备导出数据，添加序号
        const dataRows = sortedVulnerableList.map((device, index) => {
            console.log(`导出数据行 ${index+1}: 物料编号=${device.erpCode}, 机型=${device.model}`);
            return [
                index + 1, // 序号
                device.erpCode, // 物料号（ERP和物料编号是同一个东西）
                device.description, // 物料描述
                device.model, // 机型
                device.spareCount, // 建议备件数量
                device.unit, // 单位
                device.remark // 备注
            ];
        });
        
        // 创建包含标题和表头的完整数据
        const exportData = [
            ['KDZC-XXXXXX项目名称易损件清单'], // 标题行
            ['序号', '物料号', '物料描述', '机型', '建议备件数量', '单位', '备注'], // 表头行
            ...dataRows // 数据行
        ];
        
        // 创建Excel工作簿
        const worksheet = XLSX.utils.aoa_to_sheet(exportData);
        
        // 合并标题行的单元格
        worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];
        
        // 设置列宽
        worksheet['!cols'] = [
            { wch: 5 }, // 序号
            { wch: 10 }, // 物料号
            { wch: 45 }, // 物料描述
            { wch: 7 }, // 机型
            { wch: 13 }, // 建议备件数量
            { wch: 7 }, // 单位
            { wch: 5 } // 备注
        ];
        
        // 设置标题行样式
        const titleCell = worksheet['A1'];
        if (titleCell) {
            titleCell.s = {
                font: {
                    name: '等线',
                    sz: 14,
                    bold: true
                },
                alignment: {
                    horizontal: 'center',
                    vertical: 'center'
                }
            };
        }
        
        // 设置表头行样式
        for (let c = 0; c < 7; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r: 1, c });
            const cell = worksheet[cellAddr];
            if (cell) {
                cell.s = {
                    font: {
                        name: '等线',
                        sz: 11,
                        bold: true
                    },
                    alignment: {
                        horizontal: 'center',
                        vertical: 'center'
                    },
                    fill: {
                        type: 'pattern',
                        patternType: 'solid',
                        fgColor: { rgb: 'E0E0E0' } // 灰色背景
                    },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    }
                };
            }
        }
        
        // 设置数据行样式
        for (let r = 2; r <= 1 + dataRows.length; r++) {
            for (let c = 0; c < 7; c++) {
                const cellAddr = XLSX.utils.encode_cell({ r, c });
                const cell = worksheet[cellAddr];
                if (cell) {
                    // 确定对齐方式
                    let horizontalAlign = 'left';
                    if (c === 0 || c === 1 || c === 4) { // 序号、物料号、备件数
                        horizontalAlign = 'right';
                    }
                    
                    cell.s = {
                        font: {
                            name: '等线',
                            sz: 11
                        },
                        alignment: {
                            horizontal: horizontalAlign,
                            vertical: 'center'
                        },
                        border: {
                            top: { style: 'thin' },
                            bottom: { style: 'thin' },
                            left: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                }
            }
        }
        
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '易损件清单');
        
        // 生成Excel文件并下载
        XLSX.writeFile(workbook, '易损件清单.xlsx');
        alert(`易损清单导出成功！共导出 ${uniqueVulnerableList.length} 个唯一设备。`);
    } catch (error) {
        console.error('导出易损清单失败:', error);
        alert('导出易损清单失败: ' + error.message);
    }
}

// 导出黑名单
async function exportBlacklist() {
    try {
        // 从设备清单中过滤出黑名单设备
        const blacklistDevices = [];
        for (const [key, device] of devices.entries()) {
            if (device.status === '黑名单') {
                blacklistDevices.push(device);
            }
        }
        
        if (blacklistDevices.length === 0) {
            alert('没有黑名单设备可以导出');
            return;
        }
        
        // 准备导出数据，从第一行开始就是数据，没有表头
        // 第一列：物料编号，第二列：物料描述
        const dataRows = blacklistDevices.map(device => [
            device.materialId, // 物料编号
            device.description // 物料描述
        ]);
        
        // 直接使用数据行，没有标题和表头
        const exportData = [...dataRows];
        
        // 创建Excel工作簿
        const worksheet = XLSX.utils.aoa_to_sheet(exportData);
        
        // 设置列宽
        worksheet['!cols'] = [
            { wch: 15 }, // 物料编号
            { wch: 45 } // 物料描述
        ];
        
        // 设置数据行样式
        for (let r = 0; r < dataRows.length; r++) {
            for (let c = 0; c < 2; c++) {
                const cellAddr = XLSX.utils.encode_cell({ r, c });
                const cell = worksheet[cellAddr];
                if (cell) {
                    cell.s = {
                        font: {
                            name: '等线',
                            sz: 11
                        },
                        alignment: {
                            horizontal: c === 0 ? 'right' : 'left', // 物料编号靠右，物料描述靠左
                            vertical: 'center'
                        },
                        border: {
                            top: { style: 'thin' },
                            bottom: { style: 'thin' },
                            left: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                }
            }
        }
        
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '黑名单');
        
        // 生成Excel文件并下载
        XLSX.writeFile(workbook, '黑名单.xlsx');
        alert('黑名单导出成功！');
    } catch (error) {
        console.error('导出黑名单失败:', error);
        alert('导出黑名单失败: ' + error.message);
    }
}

// 导出易损件库，按机型分类生成Excel文件
async function exportVulnerableLibrary() {
    try {
        // 从设备清单中过滤出白名单设备
        const whitelistDevices = [];
        for (const [key, device] of devices.entries()) {
            if (device.status === '白名单') {
                whitelistDevices.push(device);
            }
        }
        
        if (whitelistDevices.length === 0) {
            alert('没有白名单设备可以导出');
            return;
        }
        
        // 定义需要导出的机型列表
        const models = ['系统', '摆渡车', '运输车', '砖机', '辅机'];
        
        // 按机型分类设备
        const devicesByModel = new Map();
        models.forEach(model => {
            devicesByModel.set(model, []);
        });
        
        whitelistDevices.forEach(device => {
            if (models.includes(device.model)) {
                devicesByModel.get(device.model).push(device);
            }
        });
        
        // 为每个机型创建一个Excel文件
        let successCount = 0;
        
        for (const [model, modelDevices] of devicesByModel.entries()) {
            if (modelDevices.length === 0) {
                continue;
            }
            
            // 准备导出数据，第一行是表头
            const headers = ['物料号', '物料描述', '机型', '建议备件数量', '单位', '备注'];
            
            // 数据行
            const dataRows = modelDevices.map(device => [
                device.materialId, // 物料号
                device.description, // 物料描述
                device.model, // 机型
                device.spareCount, // 建议备件数量
                device.unit, // 单位
                device.remark // 备注
            ]);
            
            // 创建包含表头和数据的完整数据
            const exportData = [
                headers, // 表头行
                ...dataRows // 数据行
            ];
            
            // 创建Excel工作簿
            const worksheet = XLSX.utils.aoa_to_sheet(exportData);
            
            // 设置列宽
            worksheet['!cols'] = [
                { wch: 15 }, // 物料号
                { wch: 45 }, // 物料描述
                { wch: 10 }, // 机型
                { wch: 15 }, // 建议备件数量
                { wch: 8 }, // 单位
                { wch: 15 } // 备注
            ];
            
            // 设置表头样式
            for (let c = 0; c < headers.length; c++) {
                const cellAddr = XLSX.utils.encode_cell({ r: 0, c });
                const cell = worksheet[cellAddr];
                if (cell) {
                    cell.s = {
                        font: {
                            name: '等线',
                            sz: 11,
                            bold: true
                        },
                        alignment: {
                            horizontal: 'center',
                            vertical: 'center'
                        },
                        fill: {
                            type: 'pattern',
                            patternType: 'solid',
                            fgColor: { rgb: 'E0E0E0' } // 灰色背景
                        },
                        border: {
                            top: { style: 'thin' },
                            bottom: { style: 'thin' },
                            left: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                }
            }
            
            // 设置数据行样式
            for (let r = 1; r < exportData.length; r++) {
                for (let c = 0; c < headers.length; c++) {
                    const cellAddr = XLSX.utils.encode_cell({ r, c });
                    const cell = worksheet[cellAddr];
                    if (cell) {
                        // 确定对齐方式
                        let horizontalAlign = 'left';
                        if (c === 0 || c === 3) { // 物料号、建议备件数量靠右对齐
                            horizontalAlign = 'right';
                        } else if (c === 2) { // 机型居中对齐
                            horizontalAlign = 'center';
                        }
                        
                        cell.s = {
                            font: {
                                name: '等线',
                                sz: 11
                            },
                            alignment: {
                                horizontal: horizontalAlign,
                                vertical: 'center'
                            },
                            border: {
                                top: { style: 'thin' },
                                bottom: { style: 'thin' },
                                left: { style: 'thin' },
                                right: { style: 'thin' }
                            }
                        };
                    }
                }
            }
            
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, model);
            
            // 生成Excel文件并下载
            XLSX.writeFile(workbook, `${model}_易损件清单.xlsx`);
            successCount++;
        }
        
        alert(`易损件库导出成功！共导出 ${successCount} 个机型的Excel文件。`);
    } catch (error) {
        console.error('导出易损件库失败:', error);
        alert('导出易损件库失败: ' + error.message);
    }
}

// 批量导入黑名单
async function importBlacklist() {
    try {
        // 选择文件
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        input.onchange = async (e) => {
            const file = e.target.files[0];
            if (!file) {
                return;
            }
            
            // 读取Excel文件
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // 解析Excel数据，从第一行开始，没有表头
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const blacklistDevices = [];
            
            // 遍历所有行，从第一行开始
            for (let r = range.s.r; r <= range.e.r; r++) {
                const materialIdCell = XLSX.utils.encode_cell({ r, c: 0 }); // 第一列：物料编号
                const descriptionCell = XLSX.utils.encode_cell({ r, c: 1 }); // 第二列：物料描述
                
                const materialId = worksheet[materialIdCell]?.v?.toString().trim();
                const description = worksheet[descriptionCell]?.v?.toString().trim();
                
                if (materialId) {
                    blacklistDevices.push({
                        materialId: materialId,
                        spareCount: 0, // 默认备件数为0
                        unit: '', // 默认单位为空
                        model: '', // 默认机型为空
                        remark: '', // 默认备注为空
                        description: description || '', // 物料描述
                        status: '黑名单' // 固定为黑名单
                    });
                }
            }
            
            if (blacklistDevices.length === 0) {
                alert('未解析到任何有效设备数据');
                return;
            }
            
            // 批量保存设备，遇到重复物料编号直接跳过
            let importedCount = 0;
            blacklistDevices.forEach(device => {
                // 确保status值有效
                device.status = '黑名单';
                
                // 直接使用materialId作为唯一键
                if (!devices.has(device.materialId)) {
                    devices.set(device.materialId, device);
                    importedCount++;
                } else {
                    console.log(`跳过重复物料编号: ${device.materialId}`);
                }
            });
            
            // 更新成功导入数量
            const successMessage = importedCount > 0 
                ? `成功导入 ${importedCount} 条黑名单设备，跳过 ${blacklistDevices.length - importedCount} 条重复设备`
                : '未导入任何设备，所有物料编号均已存在';
            
            // 保存设备清单
            saveDevices();
            
            // 重新渲染设备表格
            renderDeviceTable();
            
            alert(successMessage);
        };
        input.click();
    } catch (error) {
        console.error('批量导入黑名单失败:', error);
        alert('导入失败: ' + error.message);
    }
}

// 批量导入白名单
async function importWhitelist() {
    try {
        // 选择文件
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        input.onchange = async (e) => {
            const file = e.target.files[0];
            if (!file) {
                return;
            }
            
            // 读取Excel文件
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // 解析Excel数据，自动根据表头识别列
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const whitelistDevices = [];
            
            // 先解析表头行，获取列映射
            const headerRow = range.s.r; // 表头行号
            const headerMap = new Map(); // 表头映射，key: 表头内容, value: 列索引
            
            // 遍历表头行的所有列，建立映射关系
            for (let c = range.s.c; c <= range.e.c; c++) {
                const headerCell = XLSX.utils.encode_cell({ r: headerRow, c });
                const headerValue = worksheet[headerCell]?.v?.toString().trim();
                if (headerValue) {
                    headerMap.set(headerValue, c);
                }
            }
            
            // 定义期望的表头映射关系
            const expectedHeaders = {
                '物料号': 'materialId',
                '物料描述': 'description',
                '状态': 'status',
                '机型': 'model',
                '备件数': 'spareCount',
                '单位': 'unit',
                '备注': 'remark',
                // 支持别名
                '物料编号': 'materialId',
                '建议备件数量': 'spareCount'
            };
            
            // 构建列索引映射
            const columnMap = {};
            headerMap.forEach((colIndex, headerName) => {
                for (const [expectedHeader, fieldName] of Object.entries(expectedHeaders)) {
                    // 模糊匹配，忽略大小写
                    if (headerName.includes(expectedHeader)) {
                        columnMap[fieldName] = colIndex;
                        break;
                    }
                }
            });
            
            // 遍历所有数据行，从表头行下一行开始
            for (let r = headerRow + 1; r <= range.e.r; r++) {
                // 获取各字段值
                const materialIdCell = XLSX.utils.encode_cell({ r, c: columnMap.materialId });
                const descriptionCell = XLSX.utils.encode_cell({ r, c: columnMap.description });
                const spareCountCell = XLSX.utils.encode_cell({ r, c: columnMap.spareCount });
                const unitCell = XLSX.utils.encode_cell({ r, c: columnMap.unit });
                const modelCell = XLSX.utils.encode_cell({ r, c: columnMap.model });
                const remarkCell = XLSX.utils.encode_cell({ r, c: columnMap.remark });
                const statusCell = XLSX.utils.encode_cell({ r, c: columnMap.status });
                
                const materialId = worksheet[materialIdCell]?.v?.toString().trim();
                if (!materialId) {
                    continue;
                }
                
                const description = worksheet[descriptionCell]?.v?.toString().trim();
                const spareCount = parseInt(worksheet[spareCountCell]?.v) || 0;
                const unit = worksheet[unitCell]?.v?.toString().trim();
                const model = worksheet[modelCell]?.v?.toString().trim();
                const remark = worksheet[remarkCell]?.v?.toString().trim();
                const status = worksheet[statusCell]?.v?.toString().trim() || '白名单';
                
                whitelistDevices.push({
                    materialId: materialId,
                    description: description || '',
                    spareCount: spareCount,
                    unit: unit || '',
                    model: model || '',
                    remark: remark || '',
                    status: status
                });
            }
            
            if (whitelistDevices.length === 0) {
                alert('未解析到任何有效设备数据');
                return;
            }
            
            // 批量保存设备，遇到重复物料编号直接跳过
            let importedCount = 0;
            whitelistDevices.forEach(device => {
                // 确保status值有效
                if (!['白名单', '黑名单'].includes(device.status)) {
                    device.status = '白名单'; // 默认使用白名单
                }
                
                // 直接使用materialId作为唯一键
                if (!devices.has(device.materialId)) {
                    devices.set(device.materialId, device);
                    importedCount++;
                } else {
                    console.log(`跳过重复物料编号: ${device.materialId}`);
                }
            });
            
            // 保存设备清单
            saveDevices();
            
            // 重新渲染设备表格
            renderDeviceTable();
            
            // 更新成功导入数量
            const successMessage = importedCount > 0 
                ? `成功导入 ${importedCount} 条白名单设备，跳过 ${whitelistDevices.length - importedCount} 条重复设备`
                : '未导入任何设备，所有物料编号均已存在';
            
            alert(successMessage);
        };
        input.click();
    } catch (error) {
        console.error('批量导入白名单失败:', error);
        alert('导入失败: ' + error.message);
    }
}

// 辅助函数：按指定键去重
function uniqueByKey(array, keys) {
    const seen = new Set();
    return array.filter(item => {
        const key = keys.map(k => item[k]).join('|');
        if (seen.has(key)) {
            return false;
        }
        seen.add(key);
        return true;
    });
}

// 初始化应用
initApp();

// 事件监听（已在initApp函数中添加，此处代码冗余，建议删除）
// 如需保留，应使用dom.xxx形式：
/*
dom.tabs.forEach(tab => {
    tab.addEventListener('click', () => {
        switchTab(tab.dataset.tab);
    });
});

dom.addDeviceBtn.addEventListener('click', addDevice);
dom.saveDevicesBtn.addEventListener('click', saveDevices);
dom.loadDevicesBtn.addEventListener('click', loadDevices);
dom.selectFolderBtn.addEventListener('click', selectFiles);
dom.processFilesBtn.addEventListener('click', processFiles);
dom.exportListBtn.addEventListener('click', exportList);

dom.deviceTypeFilter.addEventListener('change', renderDeviceTable);
dom.materialIdSearch.addEventListener('input', renderDeviceTable);
*/

// 绑定导出按钮事件
document.getElementById('exportBlacklist').addEventListener('click', exportBlacklist);
document.getElementById('exportVulnerableLibrary').addEventListener('click', exportVulnerableLibrary);

// 绑定导入按钮事件
document.getElementById('importBlacklist').addEventListener('click', importBlacklist);
document.getElementById('importWhitelist').addEventListener('click', importWhitelist);
