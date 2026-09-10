# Cindy Outlook 插件

独立连接 Hotmail / Outlook.com / Live。只申请邮件权限，授权在 Cindy 里完成。

当前版本 **1.2.0**，作者 **GT**。

## 能做什么

- 搜索邮件、查看最近收件箱
- 阅读正文
- 列出文件夹
- 保存草稿、发送邮件（需明确说发送）
- 删除邮件（需明确说删除；一般进入「已删除邮件」，已在回收站则会永久删除）

## 安装

1. 下载安装包：https://github.com/guyi-dit/cindy-outlook/releases/latest/download/microsoft-outlook.cindy
2. 打开 Cindy → + 添加插件 → 安装插件...
3. 选中下载的 cindy 文件
4. 打开插件详情页，点「连接账号」，用 Hotmail 登录

源码在 plugin 目录。

## 仓库结构

- plugin/：插件源码（ghost.json、main.js、设置页、多语言）
- dist/microsoft-outlook.cindy：Cindy 安装包
