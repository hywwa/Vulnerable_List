const fs = require('fs');
const path = require('path');

// 创建dist目录（如果不存在）
const distDir = path.join(__dirname, 'dist');
if (!fs.existsSync(distDir)) {
    fs.mkdirSync(distDir);
}

// 复制静态文件
const staticFiles = [
    'server.js',
    'dbConfig.json',
    'ecosystem.config.js'
];

staticFiles.forEach(file => {
    const srcPath = path.join(__dirname, file);
    const destPath = path.join(distDir, file);
    
    try {
        fs.copyFileSync(srcPath, destPath);
        console.log(`成功复制 ${file} 到 dist/`);
    } catch (error) {
        console.error(`复制 ${file} 失败:`, error.message);
        process.exit(1);
    }
});

// 生成生产版本的package.json
const packageJsonPath = path.join(__dirname, 'package.json');
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));

// 只保留生产依赖
const productionPackageJson = {
    name: packageJson.name,
    version: packageJson.version,
    description: packageJson.description,
    main: packageJson.main,
    scripts: {
        "start": "node server.js",
        "pm2-start": "pm2 start ecosystem.config.js",
        "pm2-stop": "pm2 stop ecosystem.config.js",
        "pm2-restart": "pm2 restart ecosystem.config.js",
        "pm2-logs": "pm2 logs"
    },
    dependencies: packageJson.dependencies
};

// 写入生产版本package.json
const destPackageJsonPath = path.join(distDir, 'package.json');
fs.writeFileSync(destPackageJsonPath, JSON.stringify(productionPackageJson, null, 2));
console.log('成功生成生产版本package.json到dist/');

console.log('所有文件复制完成！');