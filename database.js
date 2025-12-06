const mysql = require('mysql2/promise');
let pool;
let deviceCache = null;
let cacheLastUpdated = null;
const CACHE_EXPIRY = 300000; // 缓存过期时间：5分钟，减少数据库查询频率

// 初始化数据库连接
function initDatabase() {
    try {
        const dbConfig = require('./dbConfig.json');
        
        // 使用更合理的连接池配置
        pool = mysql.createPool({
            ...dbConfig,
            waitForConnections: true,
            connectionLimit: 15, // 适中的连接池大小，平衡并发和资源消耗
            queueLimit: 50, // 合理的队列限制
            connectTimeout: 5000, // 连接超时时间
            idleTimeout: 30000 // 空闲连接超时
        });
        
        console.log('数据库连接初始化成功');
    } catch (error) {
        console.error('无法连接到MySQL数据库:', error.message);
        throw new Error('数据库连接失败: ' + error.message);
    }
}

// 清除设备缓存
function clearDeviceCache() {
    deviceCache = null;
    cacheLastUpdated = null;
    console.log('设备缓存已清除');
}

// 获取所有设备 - 优化版，支持同一物料编号在不同机型下的不同记录
async function getAllDevices() {
    // 检查缓存是否有效
    const now = Date.now();
    if (deviceCache && cacheLastUpdated && (now - cacheLastUpdated < CACHE_EXPIRY)) {
        console.log('使用设备缓存，共', deviceCache.size, '条记录');
        return new Map(deviceCache);
    }
    
    // 使用MySQL数据库
    try {
        // 只查询需要的字段，减少数据传输量
        const [rows] = await pool.execute('SELECT material_id, spare_count, unit, model, remark, description, status FROM devices');
        
        const devices = new Map();
        rows.forEach(row => {
            // 使用material_id + model作为复合键，支持同一物料编号在不同机型下的不同记录
            const compositeKey = `${row.material_id}|${row.model}`;
            // 直接使用字段名映射，避免不必要的转换
            devices.set(compositeKey, {
                materialId: row.material_id,
                spareCount: row.spare_count,
                unit: row.unit,
                model: row.model,
                remark: row.remark,
                description: row.description,
                status: row.status
            });
        });
        
        // 更新缓存
        deviceCache = new Map(devices);
        cacheLastUpdated = now;
        console.log('更新设备缓存，共', devices.size, '条记录');
        
        return devices;
    } catch (error) {
        console.error('获取设备清单失败:', error);
        return new Map();
    }
}

// 高效比对函数 - 核心优化，支持同一物料编号在不同机型下的不同记录
async function matchDevices(materialIdModelPairs) {
    if (!materialIdModelPairs || materialIdModelPairs.length === 0) {
        return { matched: new Map(), unmatched: [] };
    }
    
    // 提取所有物料编号，用于数据库查询
    const uniqueMaterialIds = [...new Set(materialIdModelPairs.map(pair => pair.materialId))];
    
    // 缓存键映射，用于快速查找
    const pairMap = new Map();
    materialIdModelPairs.forEach(pair => {
        const key = `${pair.materialId}|${pair.model}`;
        pairMap.set(key, pair);
    });
    
    // 检查缓存是否包含所有请求的设备
    const now = Date.now();
    if (deviceCache && cacheLastUpdated && (now - cacheLastUpdated < CACHE_EXPIRY)) {
        const matched = new Map();
        const unmatched = [];
        
        // 优先从缓存获取，O(1)复杂度
        pairMap.forEach((pair, key) => {
            const device = deviceCache.get(key);
            if (device) {
                matched.set(key, device);
            } else {
                unmatched.push(pair);
            }
        });
        
        console.log('从缓存获取匹配结果，匹配:', matched.size, '个，未匹配:', unmatched.length, '个');
        return { matched, unmatched };
    }
    
    // 缓存失效或不存在，从数据库查询
    try {
        // 使用IN语句批量查询，减少数据库往返
        // 只查询需要的字段
        const [rows] = await pool.execute(
            `SELECT material_id, spare_count, unit, model, remark, description, status 
             FROM devices 
             WHERE material_id IN (?)`,
            [uniqueMaterialIds]
        );
        
        const matched = new Map();
        const matchedKeys = new Set();
        
        // 构建匹配结果
        rows.forEach(row => {
            const device = {
                materialId: row.material_id,
                spareCount: row.spare_count,
                unit: row.unit,
                model: row.model,
                remark: row.remark,
                description: row.description,
                status: row.status
            };
            const key = `${row.material_id}|${row.model}`;
            matched.set(key, device);
            matchedKeys.add(key);
        });
        
        // 找出未匹配的物料编号-机型对
        const unmatched = [];
        pairMap.forEach((pair, key) => {
            if (!matchedKeys.has(key)) {
                unmatched.push(pair);
            }
        });
        
        console.log('从数据库获取匹配结果，匹配:', matched.size, '个，未匹配:', unmatched.length, '个');
        
        // 如果缓存为空，更新缓存，减少后续查询
        if (!deviceCache) {
            // 获取所有设备更新缓存，避免频繁查询
            console.log('缓存为空，更新全量缓存');
            await getAllDevices();
        }
        
        return { matched, unmatched };
    } catch (error) {
        console.error('批量匹配设备失败:', error);
        return { matched: new Map(), unmatched: materialIdModelPairs };
    }
}

// 添加或更新设备
async function saveDevice(device) {
    // 使用MySQL数据库
    try {
        await pool.execute(
            `INSERT INTO devices (material_id, spare_count, unit, model, remark, description, status) 
             VALUES (?, ?, ?, ?, ?, ?, ?) 
             ON DUPLICATE KEY UPDATE 
             spare_count = VALUES(spare_count), 
             unit = VALUES(unit), 
             model = VALUES(model), 
             remark = VALUES(remark), 
             description = VALUES(description),
             status = VALUES(status)`,
            [device.materialId, device.spareCount, device.unit, device.model, device.remark, device.description || null, device.status]
        );
        
        // 清除缓存，确保数据一致性
        clearDeviceCache();
        return true;
    } catch (error) {
        console.error('保存设备到数据库失败:', error);
        return false;
    }
}

// 批量保存设备 - 优化版
async function saveDevices(devices) {
    // 使用MySQL数据库
    try {
        const devicesArray = Array.from(devices.values());
        if (devicesArray.length === 0) {
            return true;
        }
        
        // 批量插入优化：使用事务和批量插入语句
        const connection = await pool.getConnection();
        
        try {
            await connection.beginTransaction();
            
            // 使用mysql2的批量插入功能，减少网络往返
            await connection.query(
                `INSERT INTO devices (material_id, spare_count, unit, model, remark, description, status) 
                 VALUES ? 
                 ON DUPLICATE KEY UPDATE 
                 spare_count = VALUES(spare_count), 
                 unit = VALUES(unit), 
                 model = VALUES(model), 
                 remark = VALUES(remark), 
                 description = VALUES(description),
                 status = VALUES(status)`,
                [devicesArray.map(device => [
                    device.materialId, 
                    device.spareCount, 
                    device.unit, 
                    device.model, 
                    device.remark, 
                    device.description || null, 
                    device.status
                ])]
            );
            
            await connection.commit();
            
            // 清除缓存
            clearDeviceCache();
            return true;
        } catch (error) {
            await connection.rollback();
            throw error;
        } finally {
            connection.release();
        }
    } catch (error) {
        console.error('批量保存设备到数据库失败:', error);
        return false;
    }
}

// 删除设备
async function deleteDevice(materialId) {
    // 使用MySQL数据库
    try {
        await pool.execute('DELETE FROM devices WHERE material_id = ?', [materialId]);
        
        // 清除缓存
        clearDeviceCache();
        return true;
    } catch (error) {
        console.error('删除设备从数据库失败:', error);
        return false;
    }
}

// 批量删除设备 - 优化版
async function deleteDevices(materialIds) {
    if (!materialIds || materialIds.length === 0) {
        return true;
    }
    
    try {
        // 批量删除，减少数据库往返
        await pool.execute(
            `DELETE FROM devices WHERE material_id IN (?)`,
            [materialIds]
        );
        
        // 清除缓存
        clearDeviceCache();
        return true;
    } catch (error) {
        console.error('批量删除设备失败:', error);
        return false;
    }
}

// 导出优化后的API
module.exports = {
    initDatabase,
    getAllDevices,
    matchDevices, // 替换getDevicesBatch为更高效的matchDevices
    saveDevice,
    saveDevices,
    deleteDevice,
    deleteDevices,
    clearDeviceCache
};