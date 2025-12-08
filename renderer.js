const { ipcRenderer } = require('electron');
const XLSX = require('xlsx-js-style');
const fs = require('fs-extra');
const path = require('path');
const db = require('./database');

// 全局变量
let devices = new Map(); // 设备清单，key: 物料编号
let selectedFolder = null;
let unknownDevices = []; // 未识别的设备
let vulnerableList = []; // 易损清单结果
let matchedDevices = []; // 白名单已匹配设备
let allFiles = []; // 处理的所有文件

// 初始化数据库连接
db.initDatabase();

// DOM元素
const tabs = document.querySelectorAll('.tab');
const tabContents = document.querySelectorAll('.tab-content');
const materialIdInput = document.getElementById('materialId');
const descriptionInput = document.getElementById('description');
const spareCountInput = document.getElementById('spareCount');
const unitInput = document.getElementById('unit');
const modelInput = document.getElementById('model');
const remarkInput = document.getElementById('remark');
const statusSelect = document.getElementById('status');
const deviceTypeFilter = document.getElementById('deviceTypeFilter');
const materialIdSearch = document.getElementById('materialIdSearch');
const addDeviceBtn = document.getElementById('addDevice');
const saveDevicesBtn = document.getElementById('saveDevices');
const loadDevicesBtn = document.getElementById('loadDevices');
const deviceTableBody = document.getElementById('deviceTableBody');
const selectFolderBtn = document.getElementById('selectFolder');
const processFilesBtn = document.getElementById('processFiles');
const exportListBtn = document.getElementById('exportList');
const fileInfoDiv = document.getElementById('fileInfo');
const summaryDiv = document.getElementById('summary');
const matchedSection = document.getElementById('matchedSection');
const matchedTableBody = document.getElementById('matchedTableBody');
const unknownSection = document.getElementById('unknownSection');
const unknownTableBody = document.getElementById('unknownTableBody');
const resultSection = document.getElementById('resultSection');
const resultTableBody = document.getElementById('resultTableBody');

// 标签页切换
function switchTab(tabName) {
    tabs.forEach(tab => tab.classList.remove('active'));
    tabContents.forEach(content => content.classList.remove('active'));
    
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
    document.getElementById(tabName).classList.add('active');
}

tabs.forEach(tab => {
    tab.addEventListener('click', () => {
        switchTab(tab.dataset.tab);
    });
});

// 设备清单维护
function addDevice() {
    const materialId = materialIdInput.value.trim();
    const description = descriptionInput.value.trim();
    const spareCount = parseInt(spareCountInput.value) || 0;
    const unit = unitInput.value.trim();
    const model = modelInput.value.trim();
    const remark = remarkInput.value.trim();
    const status = statusSelect.value;
    
    if (!materialId) {
        alert('物料编号不能为空');
        return;
    }
    
    if (!model) {
        alert('机型不能为空');
        return;
    }
    
    // 使用materialId+model作为复合键
    const compositeKey = `${materialId}|${model}`;
    
    devices.set(compositeKey, {
        materialId,
        description,
        spareCount,
        unit,
        model,
        remark,
        status
    });
    
    // 清空输入框
    materialIdInput.value = '';
    descriptionInput.value = '';
    spareCountInput.value = '';
    unitInput.value = '';
    modelInput.value = '';
    remarkInput.value = '';
    
    renderDeviceTable();
}

function renderDeviceTable() {
    // 获取当前过滤条件
    const filterType = deviceTypeFilter.value;
    const searchTerm = materialIdSearch.value.trim().toLowerCase();
    
    // 过滤设备
    const filteredDevices = [];
    for (const [materialId, device] of devices) {
        const matchesType = filterType === 'all' || device.status === filterType;
        const matchesSearch = !searchTerm || materialId.toLowerCase().includes(searchTerm);
        
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
        // 白名单和全部设备显示完整8列表头
        thead.innerHTML = `
            <tr>
                <th>物料编号</th>
                <th>物料描述</th>
                <th>备件数</th>
                <th>单位</th>
                <th>机型</th>
                <th>备注</th>
                <th>状态</th>
                <th>操作</th>
            </tr>
        `;
    }
    
    // 清空表格内容
    const tbody = deviceTable.querySelector('tbody');
    tbody.innerHTML = '';
    
    // 渲染白名单设备
    if (whiteListDevices.length > 0) {
        whiteListDevices.forEach(device => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${device.materialId}</td>
                <td>${device.description || ''}</td>
                <td>${device.spareCount}</td>
                <td>${device.unit}</td>
                <td>${device.model}</td>
                <td>${device.remark}</td>
                <td>
                    <span class="status-badge status-white">
                        ${device.status}
                    </span>
                </td>
                <td>
                    <button onclick="removeDevice('${device.materialId}')" class="danger" style="padding: 5px 10px; font-size: 12px; margin-right: 5px;">
                        删除
                    </button>
                    <button onclick="toggleDeviceStatus('${device.materialId}')" class="secondary" style="padding: 5px 10px; font-size: 12px;">
                        改为黑名单
                    </button>
                </td>
            `;
            tbody.appendChild(row);
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
                    <button onclick="removeDevice('${device.materialId}')" class="danger" style="padding: 5px 10px; font-size: 12px; margin-right: 5px;">
                        删除
                    </button>
                    <button onclick="toggleDeviceStatus('${device.materialId}')" class="secondary" style="padding: 5px 10px; font-size: 12px;">
                        改为白名单
                    </button>
                </td>
            `;
            tbody.appendChild(row);
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
        const colSpan = filterType === '黑名单' ? 4 : 8;
        emptyRow.innerHTML = `
            <td colspan="${colSpan}" style="text-align: center; padding: 20px; color: #999;">
                ${emptyMessage}
            </td>
        `;
        tbody.appendChild(emptyRow);
    }
}

async function removeDevice(materialId) {
    try {
        await db.deleteDevice(materialId);
        devices.delete(materialId);
        renderDeviceTable();
    } catch (error) {
        console.error('删除设备失败:', error);
        alert('删除设备失败: ' + error.message);
    }
}

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
        
        // 更新数据库
        await db.saveDevices(devices);
        
        // 重新渲染设备表格
        renderDeviceTable();
        
        alert(`设备已从${device.status}改为${newStatus}`);
    } catch (error) {
        console.error('切换设备状态失败:', error);
        alert('切换设备状态失败: ' + error.message);
    }
}

async function saveDevices() {
    try {
        const success = await db.saveDevices(devices);
        if (success) {
            alert('设备清单保存成功');
        } else {
            alert('设备清单保存失败');
        }
    } catch (error) {
        console.error('保存设备清单失败:', error);
        alert('保存设备清单失败: ' + error.message);
    }
}

async function loadDevices() {
    try {
        devices = await db.getAllDevices();
        renderDeviceTable();
        alert('设备清单加载成功');
    } catch (error) {
        console.error('加载设备清单失败:', error);
        alert('加载设备清单失败: ' + error.message);
    }
}

// 选择项目文件夹
async function selectFolder() {
    try {
        selectedFolder = await ipcRenderer.invoke('select-directory');
        if (selectedFolder) {
            // 获取所有Excel文件
            allFiles = await fs.readdir(selectedFolder);
            allFiles = allFiles.filter(file => 
                ['.xlsx', '.xls'].includes(path.extname(file).toLowerCase())
            );
            
            fileInfoDiv.innerHTML = `<strong>已选择文件夹:</strong> ${selectedFolder}<br>
                                     <strong>找到Excel文件:</strong> ${allFiles.length} 个<br>
                                     ${allFiles.map(file => `• ${file}`).join('<br>')}`;
            fileInfoDiv.style.display = 'block';
        }
    } catch (error) {
        console.error('选择文件夹失败:', error);
        alert('选择文件夹失败: ' + error.message);
    }
}

// 处理Excel文件
async function processFiles() {
    if (!selectedFolder || allFiles.length === 0) {
        alert('请先选择包含Excel文件的项目文件夹');
        return;
    }
    
    // 确保设备清单已加载
    if (devices.size === 0) {
        devices = await db.getAllDevices();
        if (devices.size === 0) {
            alert('请先维护设备清单');
            return;
        }
    }
    
    // 重置结果数组
    unknownDevices = [];
    vulnerableList = [];
    matchedDevices = [];
    
    let totalFiles = allFiles.length;
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
    
    // 并发控制：限制同时处理的文件数量不超过10个
    const CONCURRENCY_LIMIT = 10;
    const processResults = [];
    
    // 并发处理函数
    async function processFilesConcurrently() {
        for (let i = 0; i < allFiles.length; i += CONCURRENCY_LIMIT) {
            const batch = allFiles.slice(i, i + CONCURRENCY_LIMIT);
            const batchPromises = batch.map(async (file) => {
                const filePath = path.join(selectedFolder, file);
                try {
                    // 读取Excel文件
                    const workbook = XLSX.readFile(filePath, { cellDates: true, cellText: false });
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
                        console.log(`文件 ${file} 中未找到ERP列`);
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
                    const fileName = file.toLowerCase();
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
                        // 使用materialId+model作为复合键查找设备
                        const compositeKey = `${materialId}|${model}`;
                        const device = devices.get(compositeKey);
                        
                        if (device) {
                            // 设备已存在且机型匹配 - 先检查黑名单，再检查白名单
                            if (device.status === '黑名单') {
                                // 黑名单优先级更高，直接标记为匹配黑名单
                                fileResults.matchedBlack++;
                            } else if (device.status === '白名单') {
                                // 白名单设备，添加到易损清单
                                fileResults.vulnerable.push({
                                    erpCode: materialId,
                                    ...device,
                                    model: model // 使用根据文件名判断的机型，而不是设备清单中的机型
                                });
                                fileResults.matchedWhite++;
                            }
                        } else {
                            // 检查是否存在相同materialId但不同model的设备
                            let hasDifferentModel = false;
                            for (const [key, dev] of devices.entries()) {
                                if (dev.materialId === materialId && dev.model !== model) {
                                    hasDifferentModel = true;
                                    break;
                                }
                            }
                            
                            if (hasDifferentModel) {
                                // 存在相同materialId但不同model的设备，标记为匹配黑名单
                                fileResults.matchedBlack++;
                            } else {
                                // 完全未匹配，添加到未知设备
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
                        }
                    });
                    
                    return fileResults;
                } catch (error) {
                    console.error(`处理文件 ${file} 失败:`, error);
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
    results.forEach(result => {
        if (result) {
            vulnerableList.push(...result.vulnerable);
            unknownDevices.push(...result.unknown);
            matchedWhite += result.matchedWhite;
            matchedBlack += result.matchedBlack;
            unmatched += result.unmatched;
        }
    });
    
    // 去重处理 - 按照机型+ERP编号组合去重，同一机型内的相同ERP编号只保留一个
    unknownDevices = uniqueByKey(unknownDevices, ['model', 'erpCode']);
    vulnerableList = uniqueByKey(vulnerableList, ['model', 'erpCode']);
    
    // 处理白名单匹配设备 - 从易损清单中提取，只显示4列
    matchedDevices = vulnerableList.map(device => ({
        erpCode: device.erpCode,
        description: device.description,
        model: device.model
    }));
    
    // 更新统计信息
    summaryDiv.innerHTML = `
        <strong>处理结果:</strong><br>
        - 总文件数: ${totalFiles} 个<br>
        - 总设备数: ${totalDevices} 个<br>
        - 匹配白名单: ${matchedWhite} 个<br>
        - 匹配黑名单: ${matchedBlack} 个<br>
        - 已匹配白名单设备(去重后): ${matchedDevices.length} 个<br>
        - 未识别设备: ${unknownDevices.length} 个 (去重后)<br>
    `;
    summaryDiv.style.display = 'block';
    
    // 显示白名单匹配设备
    if (matchedDevices.length > 0) {
        matchedSection.style.display = 'block';
        renderMatchedTable();
    } else {
        matchedSection.style.display = 'none';
    }
    
    // 显示未识别设备
    if (unknownDevices.length > 0) {
        unknownSection.style.display = 'block';
        renderUnknownTable();
    } else {
        unknownSection.style.display = 'none';
        renderResultTable();
        exportListBtn.disabled = false;
    }
    
    resultSection.style.display = 'block';
}

function renderUnknownTable() {
    unknownTableBody.innerHTML = '';
    
    // 添加多选样式
    const style = document.createElement('style');
    style.textContent = `
        /* 增强多选下拉框的选中效果 */
        select[multiple] option:checked {
            background-color: #3498db !important;
            color: white !important;
        }
        select[multiple] {
            font-size: 14px;
            font-family: Arial, sans-serif;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 5px;
            outline: none;
            cursor: pointer;
        }
        select[multiple]:focus {
            border-color: #3498db;
            box-shadow: 0 0 0 2px rgba(52, 152, 219, 0.2);
        }
    `;
    document.head.appendChild(style);
    
    unknownDevices.forEach((device, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${device.erpCode}</td>
            <td><input type="text" id="materialId_${index}" value="${device.description || device.erpCode}" placeholder="物料描述" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 300px; min-width: 200px;"></td>
            <td><input type="number" id="spareCount_${index}" value="${device.spareCount}" min="0" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 80px;"></td>
            <td><input type="text" id="unit_${index}" value="${device.unit}" placeholder="单位" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 80px;"></td>
            <td>
                <select id="model_${index}" multiple="multiple" size="5" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 150px; height: auto;">
                    <option value="系统" ${device.model === '系统' ? 'selected' : ''}>系统</option>
                    <option value="摆渡车" ${device.model === '摆渡车' ? 'selected' : ''}>摆渡车</option>
                    <option value="运输车" ${device.model === '运输车' ? 'selected' : ''}>运输车</option>
                    <option value="砖机" ${device.model === '砖机' ? 'selected' : ''}>砖机</option>
                    <option value="辅机" ${device.model === '辅机' ? 'selected' : ''}>辅机</option>
                </select>
                <small style="display: ${device.isVulnerable ? 'block' : 'none'}; color: #666; margin-top: 5px;">可按住Ctrl键多选</small>
            </td>
            <td><input type="text" id="remark_${index}" value="${device.remark}" placeholder="备注" style="display: ${device.isVulnerable ? 'block' : 'none'}; width: 120px;"></td>
            <td>
                <select id="isVulnerable_${index}" onchange="toggleDeviceInfo(${index}, this.value)" style="width: 80px;">
                    <option value="true" ${device.isVulnerable ? 'selected' : ''}>是</option>
                    <option value="false" ${!device.isVulnerable ? 'selected' : ''}>否</option>
                </select>
            </td>
        `;
        unknownTableBody.appendChild(row);
    });
    
    // 添加确认按钮
    const confirmRow = document.createElement('tr');
    confirmRow.innerHTML = `
        <td colspan="7" style="text-align: center;">
            <button onclick="confirmUnknownDevices()" style="margin-top: 10px;">
                确认选择
            </button>
        </td>
    `;
    unknownTableBody.appendChild(confirmRow);
}

// 切换设备信息输入框显示
function toggleDeviceInfo(index, isVulnerable) {
    const show = isVulnerable === 'true';
    document.getElementById(`materialId_${index}`).style.display = show ? 'block' : 'none';
    document.getElementById(`spareCount_${index}`).style.display = show ? 'block' : 'none';
    document.getElementById(`unit_${index}`).style.display = show ? 'block' : 'none';
    document.getElementById(`model_${index}`).style.display = show ? 'block' : 'none';
    document.getElementById(`remark_${index}`).style.display = show ? 'block' : 'none';
}

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
        
        // 获取多选的机型
        const modelSelect = document.getElementById(`model_${i}`);
        const selectedModels = Array.from(modelSelect.selectedOptions).map(option => option.value);
        
        // 如果没有选择机型，使用默认机型
        if (selectedModels.length === 0) {
            selectedModels.push(device.model || '');
        }
        
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
            if (selectedModels.some(model => !model)) {
                alert(`第${i+1}行：机型不能为空`);
                return;
            }
        }
        
        // 保存设备原始机型（文件机型）
        const originalFileModel = device.model;
        
        // 为每个选中的机型创建一条记录
        selectedModels.forEach(selectedModel => {
            // 使用materialId+selectedModel作为复合键
            const compositeKey = `${materialId}|${selectedModel}`;
            
            // 根据是否易损创建设备对象
            const deviceObj = {
                materialId,
                description,
                status: isVulnerable ? '白名单' : '黑名单'
            };
            
            // 如果是易损，需要完整字段
            if (isVulnerable) {
                deviceObj.spareCount = spareCount;
                deviceObj.unit = unit;
                deviceObj.model = selectedModel;
                deviceObj.remark = remark;
            } else {
                // 选择否的情况：只需要编号和描述
                deviceObj.spareCount = 0;
                deviceObj.unit = '';
                deviceObj.model = selectedModel; // 保留机型，用于复合键
                deviceObj.remark = '';
            }
            
            // 添加到设备清单
            devices.set(compositeKey, deviceObj);
            
            // 如果是易损，并且机型与文件机型匹配，添加到结果
            if (isVulnerable && selectedModel === originalFileModel) {
                vulnerableList.push({
                    erpCode: device.erpCode,
                    materialId,
                    description,
                    spareCount,
                    unit,
                    model: selectedModel,
                    remark,
                    status: '白名单'
                });
            }
        });
    }
    
    // 保存更新后的设备清单
    saveDevices();
    
    // 隐藏未知设备区域
    unknownSection.style.display = 'none';
    
    // 显示结果
    renderResultTable();
    exportListBtn.disabled = false;
}

function renderMatchedTable() {
    matchedTableBody.innerHTML = '';
    
    // 更新标题，显示总数
    const matchedSection = document.getElementById('matchedSection');
    const h3 = matchedSection.querySelector('h3');
    if (h3) {
        h3.textContent = `白名单已匹配设备 (共 ${matchedDevices.length} 个)`;
    }
    
    matchedDevices.forEach(device => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${device.erpCode}</td>
            <td>${device.description}</td>
            <td>${device.model}</td>
        `;
        matchedTableBody.appendChild(row);
    });
}

function renderResultTable() {
    resultTableBody.innerHTML = '';
    
    vulnerableList.forEach(device => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${device.erpCode}</td>
            <td>${device.spareCount}</td>
            <td>${device.unit}</td>
            <td>${device.model}</td>
            <td>${device.remark}</td>
        `;
        resultTableBody.appendChild(row);
    });
}

// 导出易损清单
async function exportList() {
    if (vulnerableList.length === 0) {
        alert('没有易损设备可以导出');
        return;
    }
    
    try {
        // 准备导出数据，添加序号
        const dataRows = vulnerableList.map((device, index) => [
            index + 1, // 序号
            device.erpCode, // 物料号
            device.description, // 物料描述
            device.model, // 机型
            device.spareCount, // 建议备件数量
            device.unit, // 单位
            device.remark // 备注
        ]);
        
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
        
        // 添加单元格样式
        // 标题样式：加粗、等线、14号、居中
        // 表头样式：居中、加粗、等线、11号
        // 数据样式：等线、11号
        // 对齐方式：序号、物料号、备件数靠右，其他靠左
        
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
        
        // 生成Excel文件
        const excelBuffer = XLSX.write(workbook, { 
            bookType: 'xlsx', 
            type: 'buffer'
        });
        
        // 保存文件
        const filePath = await ipcRenderer.invoke('save-file', excelBuffer, '易损件清单.xlsx');
        if (filePath) {
            alert('易损清单导出成功！');
        }
    } catch (error) {
        console.error('导出易损清单失败:', error);
        alert('导出易损清单失败: ' + error.message);
    }
}

// 批量导入黑名单
async function importBlacklist() {
    try {
        // 选择文件 - 通过IPC调用主进程的文件选择功能
        const filePath = await ipcRenderer.invoke('select-file', 
            [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }], 
            '选择黑名单Excel文件'
        );
        
        if (!filePath) {
            return;
        }
        console.log('选择的文件:', filePath);
        
        // 读取Excel文件
        const workbook = XLSX.readFile(filePath);
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
        
        console.log('解析到的黑名单设备数量:', blacklistDevices.length);
        
        if (blacklistDevices.length === 0) {
            alert('未解析到任何有效设备数据');
            return;
        }
        
        // 批量保存到数据库
        const devicesMap = new Map();
        blacklistDevices.forEach(device => {
            devicesMap.set(device.materialId, device);
        });
        
        const saveResult = await db.saveDevices(devicesMap);
        if (saveResult) {
            alert(`成功导入 ${blacklistDevices.length} 条黑名单设备`);
            // 重新加载设备清单
            await loadDevices();
        } else {
            alert('导入失败，请重试');
        }
        
    } catch (error) {
        console.error('批量导入黑名单失败:', error);
        alert('导入失败: ' + error.message);
    }
}

// 批量导入白名单
async function importWhitelist() {
    try {
        // 选择文件 - 通过IPC调用主进程的文件选择功能
        const filePath = await ipcRenderer.invoke('select-file', 
            [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }], 
            '选择白名单Excel文件'
        );
        
        if (!filePath) {
            return;
        }
        console.log('选择的文件:', filePath);
        
        // 读取Excel文件
        const workbook = XLSX.readFile(filePath);
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
            const statusCell = columnMap.status ? XLSX.utils.encode_cell({ r, c: columnMap.status }) : null;
            const modelCell = XLSX.utils.encode_cell({ r, c: columnMap.model });
            const spareCountCell = XLSX.utils.encode_cell({ r, c: columnMap.spareCount });
            const unitCell = XLSX.utils.encode_cell({ r, c: columnMap.unit });
            const remarkCell = XLSX.utils.encode_cell({ r, c: columnMap.remark });
            
            const materialId = worksheet[materialIdCell]?.v?.toString().trim();
            const description = worksheet[descriptionCell]?.v?.toString().trim();
            const statusValue = statusCell ? worksheet[statusCell]?.v?.toString().trim() : null;
            const model = worksheet[modelCell]?.v?.toString().trim();
            const spareCount = worksheet[spareCountCell]?.v ? parseInt(worksheet[spareCountCell].v) || 0 : 0;
            const unit = worksheet[unitCell]?.v?.toString().trim();
            const remark = worksheet[remarkCell]?.v?.toString().trim();
            
            if (materialId) {
                // 如果Excel中包含状态列，使用Excel中的状态，否则默认为白名单
                let finalStatus = '白名单';
                if (statusValue) {
                    // 处理状态值，转换为标准状态
                    if (statusValue.includes('黑') || statusValue === '黑名单') {
                        finalStatus = '黑名单';
                    } else if (statusValue.includes('白') || statusValue === '白名单') {
                        finalStatus = '白名单';
                    }
                }
                
                whitelistDevices.push({
                    materialId: materialId,
                    description: description || '', // 物料描述
                    status: finalStatus, // 使用Excel中的状态或默认白名单
                    model: model || '', // 机型
                    spareCount: spareCount, // 备件数
                    unit: unit || '', // 单位
                    remark: remark || '', // 备注
                });
            }
        }
        
        console.log('解析到的白名单设备数量:', whitelistDevices.length);
        
        if (whitelistDevices.length === 0) {
            alert('未解析到任何有效设备数据');
            return;
        }
        
        // 批量保存到数据库
        const devicesMap = new Map();
        whitelistDevices.forEach(device => {
            devicesMap.set(device.materialId, device);
        });
        
        const saveResult = await db.saveDevices(devicesMap);
        if (saveResult) {
            alert(`成功导入 ${whitelistDevices.length} 条白名单设备`);
            // 重新加载设备清单
            await loadDevices();
        } else {
            alert('导入失败，请重试');
        }
        
    } catch (error) {
        console.error('批量导入白名单失败:', error);
        alert('导入失败: ' + error.message);
    }
}

// 事件监听
addDeviceBtn.addEventListener('click', addDevice);
saveDevicesBtn.addEventListener('click', saveDevices);
loadDevicesBtn.addEventListener('click', loadDevices);
selectFolderBtn.addEventListener('click', selectFolder);
processFilesBtn.addEventListener('click', processFiles);
exportListBtn.addEventListener('click', exportList);

// 添加批量导入黑名单按钮事件监听
document.getElementById('importBlacklist').addEventListener('click', importBlacklist);

// 添加批量导入白名单按钮事件监听
document.getElementById('importWhitelist').addEventListener('click', importWhitelist);

// 添加设备类型过滤事件监听
deviceTypeFilter.addEventListener('change', renderDeviceTable);

// 添加物料编号搜索事件监听
materialIdSearch.addEventListener('input', renderDeviceTable);

// 初始化
window.removeDevice = removeDevice;
window.confirmUnknownDevices = confirmUnknownDevices;
window.toggleDeviceInfo = toggleDeviceInfo;
window.toggleDeviceStatus = toggleDeviceStatus;

// 尝试加载设备清单
async function initDevices() {
    try {
        devices = await db.getAllDevices();
        renderDeviceTable();
    } catch (error) {
        console.log('没有找到设备清单，将创建新的:', error.message);
    }
}

// 去重函数 - 支持单键或多键组合去重
function uniqueByKey(arr, keys) {
    const seen = new Set();
    return arr.filter(item => {
        // 如果是单键，直接使用该键值；如果是多键，组合成字符串
        const keyValue = Array.isArray(keys) 
            ? keys.map(key => item[key]).join('|') 
            : item[keys];
        
        if (seen.has(keyValue)) {
            return false;
        }
        seen.add(keyValue);
        return true;
    });
}

// 初始化设备清单
initDevices();