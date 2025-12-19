GaiZhangYe 盖章页工具 HTML 可视化服务
=========================

## 打包模式
本项目采用 onefile 模式打包，这是最快的打包和运行方式。

## 项目结构
```
dist_web/
├── GaiZhangYe.exe        # 单个可执行文件（包含所有依赖）
└── business_data/        # 业务数据目录
```

## 运行方式
1. 确保 `business_data` 目录与 `GaiZhangYe.exe` 位于同一目录下
2. 直接双击运行 `GaiZhangYe.exe` 即可
3. 服务将在 `http://localhost:5001` 自动运行并打开浏览器

## 功能说明
- 这是一个盖章页工具的 HTML 可视化服务
- 支持多种文档处理功能
- 服务自动管理端口，会自动清除占用5001端口的旧进程

## 注意事项
- 首次运行可能需要稍作等待
- 确保系统没有防火墙或安全软件阻止端口5001
- `business_data` 目录与 `GaiZhangYe.exe` 必须位于同一目录
