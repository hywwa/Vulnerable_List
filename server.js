const express = require('express');
const mysql = require('mysql2/promise');
const cors = require('cors');
const dbConfig = require('./dbConfig.json');

const app = express();
const PORT = process.env.PORT || 3002;

// 配置中间件
app.use(cors());
app.use(express.json());

// 配置静态文件服务
// 当在dist目录外运行时，使用dist作为静态目录；当在dist目录内运行时，使用当前目录
const isInDist = __dirname.endsWith('dist');
const staticDir = isInDist ? '.' : 'dist';
app.use(express.static(staticDir));
console.log(`静态文件目录: ${staticDir}`);

// 处理favicon.ico请求，避免404错误
app.get('/favicon.ico', (req, res) => {
    res.status(204).end(); // 204 No Content
});

// 数据库连接池
let pool;

// 初始化数据库连接
async function initDatabase() {
    try {
        pool = mysql.createPool({
            ...dbConfig,
            waitForConnections: true,
            connectionLimit: 15,
            queueLimit: 50,
            connectTimeout: 5000,
            idleTimeout: 30000
        });
        
        // 测试连接
        const connection = await pool.getConnection();
        await connection.ping();
        connection.release();
        
        console.log('数据库连接初始化成功');
    } catch (error) {
        console.error('无法连接到MySQL数据库:', error.message);
        throw new Error('数据库连接失败: ' + error.message);
    }
}

// API 端点：获取所有设备
app.get('/api/devices', async (req, res) => {
    try {
        const [rows] = await pool.execute('SELECT material_id, spare_count, unit, model, remark, description, status FROM devices');
        
        // 将结果转换为客户端需要的格式
        const devices = rows.map(row => ({
            materialId: row.material_id,
            spareCount: row.spare_count,
            unit: row.unit,
            model: row.model,
            remark: row.remark,
            description: row.description,
            status: row.status
        }));
        
        res.json(devices);
    } catch (error) {
        console.error('获取设备清单失败:', error);
        res.status(500).json({ error: '获取设备清单失败: ' + error.message });
    }
});

// API 端点：保存设备
app.post('/api/devices', async (req, res) => {
    try {
        const devices = req.body;
        if (!Array.isArray(devices)) {
            return res.status(400).json({ error: '请求体必须是设备数组' });
        }
        
        if (devices.length === 0) {
            return res.json({ success: true });
        }
        
        // 批量插入或更新设备
        const connection = await pool.getConnection();
        
        try {
            await connection.beginTransaction();
            
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
                [devices.map(device => [
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
            res.json({ success: true });
        } catch (error) {
            await connection.rollback();
            throw error;
        } finally {
            connection.release();
        }
    } catch (error) {
        console.error('保存设备到数据库失败:', error);
        res.status(500).json({ error: '保存设备到数据库失败: ' + error.message });
    }
});

// API 端点：删除设备
app.delete('/api/devices/:materialId', async (req, res) => {
    try {
        const { materialId } = req.params;
        await pool.execute('DELETE FROM devices WHERE material_id = ?', [materialId]);
        res.json({ success: true });
    } catch (error) {
        console.error('删除设备从数据库失败:', error);
        res.status(500).json({ error: '删除设备从数据库失败: ' + error.message });
    }
});

// API 端点：匹配设备
app.post('/api/devices/match', async (req, res) => {
    try {
        const materialIdModelPairs = req.body;
        if (!Array.isArray(materialIdModelPairs)) {
            return res.status(400).json({ error: '请求体必须是物料编号-机型对数组' });
        }
        
        if (materialIdModelPairs.length === 0) {
            return res.json({ matched: [], unmatched: [] });
        }
        
        // 提取所有物料编号
        const uniqueMaterialIds = [...new Set(materialIdModelPairs.map(pair => pair.materialId))];
        
        // 批量查询设备
        const [rows] = await pool.execute(
            `SELECT material_id, spare_count, unit, model, remark, description, status 
             FROM devices 
             WHERE material_id IN (?)`,
            [uniqueMaterialIds]
        );
        
        // 构建设备映射
        const deviceMap = new Map();
        rows.forEach(row => {
            const key = `${row.material_id}|${row.model}`;
            deviceMap.set(key, {
                materialId: row.material_id,
                spareCount: row.spare_count,
                unit: row.unit,
                model: row.model,
                remark: row.remark,
                description: row.description,
                status: row.status
            });
        });
        
        // 匹配设备
        const matched = [];
        const unmatched = [];
        
        materialIdModelPairs.forEach(pair => {
            const key = `${pair.materialId}|${pair.model}`;
            const device = deviceMap.get(key);
            if (device) {
                matched.push(device);
            } else {
                unmatched.push(pair);
            }
        });
        
        res.json({ matched, unmatched });
    } catch (error) {
        console.error('批量匹配设备失败:', error);
        res.status(500).json({ error: '批量匹配设备失败: ' + error.message });
    }
});

// 启动服务器
async function startServer() {
    try {
        // 初始化数据库连接
        await initDatabase();
        
        // 启动 Express 服务器
        app.listen(PORT, () => {
            console.log(`服务器正在运行，端口号: ${PORT}`);
            console.log(`API 访问地址: http://localhost:${PORT}/api`);
            console.log(`前端访问地址: http://localhost:${PORT}`);
        });
    } catch (error) {
        console.error('启动服务器失败:', error);
        process.exit(1);
    }
}

// 启动服务器
startServer();
