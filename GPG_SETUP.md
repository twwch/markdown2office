# GPG 密钥设置指南

本文档说明如何为 Maven Central 发布配置 GPG 密钥。

## 1. 生成 GPG 密钥

```bash
# 生成新的 GPG 密钥对
gpg --full-generate-key

# 选择以下选项：
# - 密钥类型：RSA and RSA (默认)
# - 密钥长度：4096
# - 有效期：0 (永不过期)
# - 输入您的姓名、邮箱和注释
# - 设置一个强密码（这将是您的 GPG_PASSPHRASE）
```

## 2. 查看并导出密钥

```bash
# 列出密钥
gpg --list-secret-keys --keyid-format LONG

# 输出示例：
# sec   rsa4096/ABCDEF1234567890 2024-01-01 [SC]
#       1234567890ABCDEF1234567890ABCDEF12345678
# uid                 [ultimate] Your Name <your.email@example.com>

# 记录密钥 ID（ABCDEF1234567890 部分）
```

## 3. 导出私钥

有两种方式导出私钥，选择其中一种：

### 方式 A：直接导出（推荐）
```bash
# 导出私钥（替换 YOUR_KEY_ID 为您的实际密钥 ID）
gpg --armor --export-secret-keys YOUR_KEY_ID > private.key

# 查看导出的密钥内容
cat private.key
```

### 方式 B：Base64 编码导出
```bash
# 导出并 base64 编码
gpg --armor --export-secret-keys YOUR_KEY_ID | base64 > private-base64.key

# 查看编码后的内容
cat private-base64.key
```

## 4. 发布公钥到密钥服务器

```bash
# 发布到密钥服务器（替换 YOUR_KEY_ID）
gpg --keyserver keyserver.ubuntu.com --send-keys YOUR_KEY_ID
gpg --keyserver keys.openpgp.org --send-keys YOUR_KEY_ID
gpg --keyserver pgp.mit.edu --send-keys YOUR_KEY_ID
```

## 5. 配置 GitHub Secrets

在您的 GitHub 仓库设置中添加以下 Secrets：

1. 访问：`https://github.com/YOUR_USERNAME/markdown2office/settings/secrets/actions`

2. 添加以下 Secrets：

### GPG_PRIVATE_KEY
- 如果使用方式 A：复制 `private.key` 文件的全部内容（包括 BEGIN 和 END 行）
- 如果使用方式 B：复制 `private-base64.key` 文件的内容（单行 base64 字符串）

示例（方式 A）：
```
-----BEGIN PGP PRIVATE KEY BLOCK-----

lQdGBGV... (很长的字符串)
...
-----END PGP PRIVATE KEY BLOCK-----
```

示例（方式 B）：
```
LS0tLS1CRUdJTi... (单行 base64 字符串)
```

### GPG_PASSPHRASE
- 输入您在生成密钥时设置的密码

### MAVEN_USERNAME
- 您的 Sonatype JIRA 用户名

### MAVEN_PASSWORD
- 您的 Sonatype JIRA 密码

## 6. 验证配置

运行以下命令验证密钥是否正确：

```bash
# 在本地测试导入
echo "$YOUR_PRIVATE_KEY_CONTENT" | gpg --batch --import

# 或者如果是 base64 编码的
echo "$YOUR_BASE64_KEY_CONTENT" | base64 --decode | gpg --batch --import

# 检查是否导入成功
gpg --list-secret-keys
```

## 常见问题

### 错误：no valid OpenPGP data found
- 原因：密钥格式不正确或损坏
- 解决：重新导出密钥，确保复制完整内容

### 错误：GPG signing failed
- 原因：密码不正确或密钥未正确导入
- 解决：检查 GPG_PASSPHRASE 是否正确

### 错误：密钥服务器无法访问
- 解决：尝试不同的密钥服务器，或稍后重试

## 安全注意事项

1. **永远不要** 将私钥提交到代码仓库
2. **定期更新** GPG 密钥
3. **使用强密码** 保护您的私钥
4. **备份密钥** 到安全的地方

## 清理临时文件

完成配置后，删除本地的密钥文件：

```bash
rm -f private.key private-base64.key
```