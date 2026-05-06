---
name: fill-mabang
description: 读取桌面"马帮信息.xlsx"，使用Playwright自动填写马帮ERP订单表单。当用户说"填马帮"时自动调用。
allowed-tools: Bash(python *) Bash(taskkill*) Bash(rm *)
---

# 填写马帮ERP订单表单

读取桌面Excel数据，自动填写马帮ERP的新建订单表单。

## 执行步骤

运行脚本完成填写：

```
python <skill-directory>/scripts/mabang_fill.py
```

脚本会：
1. 读取 `~/Desktop/马帮信息.xlsx` 中的订单数据
2. 启动Playwright浏览器（使用持久化session保持登录）
3. 导航到马帮ERP新建订单页面
4. 在iframe中按正确的字段映射自动填写所有字段
5. 输出填写结果

## 注意事项

- 如果浏览器提示未登录，告诉用户在浏览器中手动登录，脚本会自动等待并继续
- 所属国家下拉框是自定义组件，无法自动化，跳过
- 订单号(platformOrderId)保持系统默认值，不要填写
- 填完后直接告诉用户"填好了"，不需要截图确认
- 浏览器保持打开，剩下的交给用户操作，无需监控
