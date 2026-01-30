# OnlyOffice DocumentServer 插件配置指南

## 插件默认状态

**插件功能默认是开启的**，无需额外配置。只需要：

1. 将插件文件放到正确的目录
2. 重启 DocumentServer（或等待自动加载）

## Docker 安装配置

### 1. 创建插件目录

在宿主机上创建插件目录：

```bash
sudo mkdir -p /opt/onlyoffice/plugins
```

### 2. 更新 Docker 运行命令

添加插件目录挂载：

```bash
docker run -i -t -d \
  -p 8101:80 \
  --restart=always \
  --name onlyoffice-document-server \
  -e JWT_ENABLED=true \
  -e JWT_SECRET=keV5IcrAl9rO98WsLes13JaQ0ENXxHKkHsvpi4LjpE4= \
  -e JWT_HEADER=Authorization \
  -v /opt/onlyoffice/logs:/var/log/onlyoffice \
  -v /opt/onlyoffice/data:/var/www/onlyoffice/Data \
  -v /opt/onlyoffice/lib:/var/lib/onlyoffice \
  -v /opt/onlyoffice/db:/var/lib/postgresql \
  -v /opt/onlyoffice/fonts:/usr/share/fonts/truetype/custom \
  -v /opt/onlyoffice/cache:/var/lib/onlyoffice/docbuilder \
  -v /opt/onlyoffice/plugins:/var/www/onlyoffice/documentserver/sdkjs-plugins \
  -v /opt/onlyoffice/etc:/etc/onlyoffice/documentserver \
  onlyoffice/documentserver
```

**关键挂载点：**

- `-v /opt/onlyoffice/plugins:/var/www/onlyoffice/documentserver/sdkjs-plugins` - 插件目录

### 3. 部署插件

将插件复制到宿主机插件目录：

```bash
# 复制 comment-helper 插件
sudo cp -r /path/to/onlyoffice-vue-demo/plugins/comment-helper /opt/onlyoffice/plugins/

# 复制 jsapi-executor 插件
sudo cp -r /path/to/onlyoffice-vue-demo/plugins/jsapi-executor /opt/onlyoffice/plugins/

# 复制 text-search 插件（如果存在）
sudo cp -r /path/to/onlyoffice-vue-demo/plugins/text-search /opt/onlyoffice/plugins/
```

### 4. 重启 DocumentServer

```bash
# 重启容器
docker restart onlyoffice-document-server

# 或查看日志确认插件加载
docker logs -f onlyoffice-document-server
```

## 验证插件是否加载

### 方法 1：检查容器内文件

```bash
# 进入容器
docker exec -it onlyoffice-document-server bash

# 检查插件目录
ls -la /var/www/onlyoffice/documentserver/sdkjs-plugins/

# 应该能看到：
# comment-helper/
# jsapi-executor/
# text-search/
```

### 方法 2：在编辑器中查看

1. 打开 OnlyOffice 编辑器
2. 查看工具栏，应该能看到插件按钮
3. 例如："批注助手"、"JsAPI 执行器" 等

### 方法 3：检查日志

```bash
# 查看 DocumentServer 日志
docker logs onlyoffice-document-server | grep -i plugin

# 或查看详细日志
tail -f /opt/onlyoffice/logs/documentserver/docservice/out.log
```

## 手动配置插件（可选）

如果需要自定义插件设置，可以修改 `local.json`：

### 1. 创建配置文件目录

```bash
sudo mkdir -p /opt/onlyoffice/etc
```

### 2. 创建 local.json（如果需要）

```bash
sudo nano /opt/onlyoffice/etc/local.json
```

内容示例：

```json
{
  "services": {
    "CoAuthoring": {
      "server": {
        "plugins": {
          "enabled": true,
          "path": "sdkjs-plugins"
        }
      }
    }
  }
}
```

**注意：** 插件默认是开启的，通常不需要修改此配置。

### 3. 重启 DocumentServer

```bash
docker restart onlyoffice-document-server
```

## 插件目录结构

每个插件应该有以下结构：

```
plugin-name/
├── config.json          # 插件配置文件（必需）
├── plugin.js            # 插件主代码（必需）
└── resources/
    └── icon.svg         # 插件图标（可选）
```

## 常见问题

### Q1: 插件没有显示在编辑器中？

**检查清单：**

1. ✅ 插件文件是否在正确目录：`/var/www/onlyoffice/documentserver/sdkjs-plugins/`
2. ✅ `config.json` 文件是否存在且格式正确
3. ✅ `plugin.js` 文件是否存在
4. ✅ 是否重启了 DocumentServer
5. ✅ 检查浏览器控制台是否有错误

### Q2: 如何查看插件加载日志？

```bash
# 查看容器日志
docker logs onlyoffice-document-server

# 或查看详细日志文件
docker exec -it onlyoffice-document-server tail -f /var/log/onlyoffice/documentserver/docservice/out.log
```

### Q3: 插件执行出错怎么办？

1. 打开浏览器开发者工具（F12）
2. 查看 Console 标签页的错误信息
3. 检查插件代码中的 `Api.GetConsole().Log()` 输出

### Q4: 如何禁用某个插件？

```bash
# 方法 1：删除插件目录
sudo rm -rf /opt/onlyoffice/plugins/plugin-name

# 方法 2：重命名插件目录（添加 .disabled 后缀）
sudo mv /opt/onlyoffice/plugins/plugin-name /opt/onlyoffice/plugins/plugin-name.disabled

# 然后重启
docker restart onlyoffice-document-server
```

## 快速部署脚本

创建 `deploy-plugins.sh`：

```bash
#!/bin/bash

# 插件源目录
PLUGIN_SOURCE="/path/to/onlyoffice-vue-demo/plugins"
# 目标目录
PLUGIN_TARGET="/opt/onlyoffice/plugins"

# 创建目标目录
sudo mkdir -p $PLUGIN_TARGET

# 复制所有插件
sudo cp -r $PLUGIN_SOURCE/* $PLUGIN_TARGET/

# 设置权限
sudo chown -R 101:101 $PLUGIN_TARGET
sudo chmod -R 755 $PLUGIN_TARGET

# 重启 DocumentServer
docker restart onlyoffice-document-server

echo "插件部署完成！"
```

使用：

```bash
chmod +x deploy-plugins.sh
./deploy-plugins.sh
```

## 总结

1. **插件默认开启**：无需额外配置
2. **只需挂载目录**：添加 `-v /opt/onlyoffice/plugins:/var/www/onlyoffice/documentserver/sdkjs-plugins`
3. **复制插件文件**：将插件复制到 `/opt/onlyoffice/plugins/`
4. **重启服务**：`docker restart onlyoffice-document-server`
5. **验证加载**：在编辑器中查看插件按钮

完成以上步骤后，插件就可以正常使用了！
