# -*- coding: utf-8 -*-
"""
设计计划时间表生成器 v4.5中国院版
=====================================================================
依据《建筑工程设计项目专业配合规定》(2015版, 中国建筑设计院有限公司)
第二节「初步设计配合详细要求」 / 第三节「施工图设计配合详细要求」，
根据用户输入的项目名称、设计编号、设计阶段与起止日期，自动生成可编辑的
Excel 设计计划时间表(.xlsx)。

表格特征：
  * 以起止时间逐日排布，一周为一行（周一～周日共 7 个日期列）；
  * 标题区显示项目名称、设计编号、设计阶段、起止日期与周期天数；
  * 各专业配合节点按其时间百分比自动换算成日期，填入对应日期格，
    并写明具体配合内容；
  * 周六/周日及法定节假日（含调休补班处理）底色为淡绿色，与工作日区分；
  * 支持「避开节假日」：将落在休息日的节点迁移到距离最近的工作日，
    并注明原日期（由X月X日（节假日）迁移而来）；
  * 内容区隔行交替底色（斑马纹），方便阅读；
  * 支持输出版本：完整版 / 简洁版 / 同时生成；简洁版仅显示节点名称与文档章节编号
    （如“建筑专业提图，内容详专业配合规定第2节2.2.1”），不显示具体内容；
  * 支持多风格输出（v4.0）：classic 经典黑白打印友好 / blue 蓝白商务 /
    green 淡绿清新 / warm 暖色简约，风格主要影响标题、表头、阶段行、边框与
    强调色配色；法定休息日淡绿底、节点重合浅红底、斑马纹等硬性偏好各风格保留；
  * 周期输入方式（v3.7）：方式1 开始日期+结束日期；方式2 开始日期+设计计划
    周期（设计总时间，天），自动推算结束日期 end = start + (periodDays - 1) 天；
  * 支持出图制作时间（天，默认 1）：最终节点拆分为“开始出图”（end-N+1）
    与“交付”（end）两个节点，交付与整体周期最终日期重合（N=1 同日不判重合）；
  * 两个及以上时间设定（百分比）不同的节点落在同一日期时，该日期表格底色
    为浅红色，并提示“由于整体周期过短，导致节点重合，需要人工协调”；
    文档中原本就设定在同一百分比（如多个 0% 节点、多个 27% 节点）的节点
    同日属正常设计，不视为重合；自定义节点不参与重合判定；
  * 生成 Excel 时弹出系统保存对话框（tkinter），由用户自行决定保存位置；
    GUI 不可用时回退为提示用户手动输入保存路径。
  * 表格下方注明编制依据（引用文档标题及对应章节）与软件作者信息。

v4.5中国院版（自 v5.3.3 裁剪）：仅保留覆写节点全部功能，取消应用标准/自定义节点/
样式界面化/配置管理/节假日管理界面/节点落地预览等自定义菜单与 CLI 分支。
    - 覆写节点：导出覆写模板、加载覆写节点（可多选叠加）、恢复初始状态；解析时收集
      非法百分比/越界/重复/空内容明细，加载后弹解析报告；生成前预检列出重合/紧密/迁移节点。
    - 保留核心流程：重合预检（_run_preflight）、内置节假日自动加载（holidays.json）、
      多语种输出（内置词典）、内置样式表（classic/blue/green/warm）。
    软件作者：JaJa+AI 　版本：v4.5中国院版
  * A1 节假日/调休数据完整性：_HOLIDAYS/_WORKDAYS 补全 2027-2030（2027 春节
    补足 7 天及对应调休补班）；2027 及以后为预测值（按天文历法+通行惯例推算），
    程序加载时按年份提示“预测值，以国务院当年公布为准”。
  * A2 出图制作时间上限校验：drawing_days ≥ 项目周期天数时拒绝并提示合理范围
    （CLI 与 GUI 提交路径均校验）。
  * A3 节假日外置：运行支持目录 holidays.json 优先加载（允许覆盖内置数据）；
    项目起止日期跨越数据未收录年份时输出警告（CLI 打印 / GUI 弹窗或状态提示）。
  * A4 取整模式：新增 --round-mode floor|ceil|round（默认 round），控制日期
    百分比换算取整方式；GUI 增加对应下拉选项。
  * L1 节点清单信息补全：新增“原定日期 / 迁移来源 / 重合状态”三列。
  * L2 周期过长提示：项目周期 >730 天时提示（GUI 确认弹窗 / CLI 与交互打印警告）。
  * L3 行高估算修正：按中文 2 单位、英文 1 单位与列宽估算换行行数，避免译文
    或长文本被截断。
  * U1 GUI 追加自定义节点：自定义菜单新增“追加自定义节点”，生成时真正传入
    主流程（main 的 gui_opts 读取 custom_nodes）。
  * U2 GUI 结果摘要：生成完成后弹出摘要（输出路径、节点总数、里程碑/出图日期、
    重合警告等），再询问是否打开生成文件。
  * U3 GUI 双版本输出：输出版本改为 完整版/简洁版/完整+简洁 三选一下拉。
  * U4 GUI 记忆参数：save_config 增加阶段、开始/结束日期，GUI 启动时预填。
  * U5 GUI 加载节假日文件：自定义菜单新增“加载节假日文件”，写入运行支持目录
    holidays.json 并提示生效范围。
  * U6 GUI 可缩放滚动：主窗口 resizable + 内容区滚动条。

v3.1 新增（面向建筑设计管理者与使用者）：
  * 「节点清单」汇总工作表：在同一 xlsx 中新增“节点清单”sheet，按日期顺序
    列出全部节点（序号/日期/星期/所属阶段/节点名称/具体内容或章节编号/里程碑标记），
    便于管理者快速核对与打印；
  * 里程碑识别与高亮：将各次专业协调会、专业间核对/会审与会签、图纸归档、
    出图等关键节点识别为里程碑，在日历主表中对应节点文本加“【里程碑】”前缀
    并加粗显示，在节点清单表中单独标记；
  * 项目统计信息区：在日历主表底部显示设计周期总天数、总周数、工作日天数、
    休息日天数、节点总数、里程碑数、计划出图日期，帮助管理者掌握项目节奏；
  * 配置记忆：程序启动时自动读取同目录 config.json（记录上次的项目名称、
    设计编号、出图制作天数、是否避开节假日、输出格式等），交互模式下预填
    默认值，用户直接回车可沿用上次配置；无配置文件时自动创建；
  * 自定义节点追加：交互模式下询问是否追加自定义节点（名称+日期+内容），
    可追加多个，合并进日历主表与节点清单表（标注为“自定义节点”）；
  * 生成后自动打开：生成成功后询问是否立即打开生成的 Excel
    （Windows 用 os.startfile，Linux 用 xdg-open）；
  * 短周期预警：交互模式下若设计周期少于 15 天，提示“周期过短，节点可能
    大量重合”并询问是否继续。

v3.2 修订（规则缺陷修正）：
  * 节点重合判定修正：上传文档中各章节本来就设定在同一时间（同一百分比，
    如多个 0% 节点、多个 27% 节点）的多个节点同日属正常设计，不再标浅红、
    不再提示“由于整体周期过短，导致节点重合，需要人工协调”；仅当同一日期
    存在两个及以上原始时间设定（百分比）不同的节点时才判定为异常重合
    （浅红底色+提示）；自定义节点不参与重合判定；
  * 最终节点合并（v3.7 起拆分为“开始出图”+“交付”，见 v3.7 修订）：将
    最后的“出图”节点与图纸归档节点合并为“最终节点：出图与图纸归档”，
    按出图制作时间 N 天设置，制作区间为 end-N+1 ~ end，末天与设计阶段
    截止日一致；表格标注在截止日日期格，内容注明制作区间，不再单独保留
    落在结束日的归档节点；
  * 阶段顺序约束：第五阶段（初步设计）/第七阶段（施工图）的“专业间核对、
    会审与会签”节点安排在最终节点（图纸归档/出图）之前，两者不重叠；
    会审日期由原“设计文件归档之前（落结束日）”改为 end-N（最终节点开始
    前一天）；该日为休息日且启用避开节假日时向前取最近工作日（不越过
    end-N）；未启用避开节假日时按原休息日日期放置。

v3.3 修订（施工图阶段归属修正）：
  * 逐条核对《建筑工程设计项目专业配合规定》第三节施工图各节点阶段归属与
    时间百分比，修正第三/第四阶段节点数据：原标注“第三阶段 第二版作业图
    阶段”但提出时间已超出该阶段周期（55%W～65%W）的“建筑资料2”
    （70%W）、“水、暖专业提建筑、结构资料”（75%W）调整为“第四阶段 管道
    综合与对图阶段”（71%W～76%W），消除“第四阶段节点日期早于第三阶段
    节点日期”的阶段顺序错误；
  * 阶段顺序校验：按百分比换算日期后，各阶段严格按第一→第八阶段顺序排列，
    后一阶段任一节点日期不再早于前一阶段节点日期（附件实测通过）。

v3.4 修订（阶段周期表达规则）：
  * 适用条件：阶段标注百分比起止周期（如 76%W～82%W、xx%P～yy%P），且该
    阶段内具体内容不包含任何独立精确时间节点（内容仅注明“本阶段 xx%~yy%
    内完成”）；
  * 处理规则：符合条件的阶段，其全部内容统一在阶段周期开始百分比换算日期
    的表格内表达；在阶段周期结束百分比换算日期的表格内标注“X阶段结束”
    （如“第五阶段（提清资料）结束”，表内显示为“▷ 阶段结束标注”，节点
    清单显示为“【阶段结束】”行）；
  * 适用阶段核对：施工图第五阶段“提清资料阶段”（76%W～82%W）的水、暖、
    电、智能化四个提清资料节点符合条件（仅注明本阶段内完成、无独立时间
    点），按上述规则表达；初步设计 6 个阶段各节点均有独立百分比标注，第五
    阶段会审与会签、第六阶段图纸归档为“设计文件归档/交付之前”非百分比
    标注，均不适用本规则，保持原表达；
  * 阶段结束标注处理：参与“避开节假日”迁移（落休息日时迁移至最近工作日
    并注明原日期），保证所有日期标记在工作日表达、语义不丢失；不参与节点
    重合判定，不计入节点总数与里程碑统计；
  * 兼容性：v3.2 的会审节点前移、同百分比不标红等规则保持不变（最终节点
    合并自 v3.7 起改为拆分“开始出图”+“交付”，见 v3.7 修订）。

v3.6 修订（最后一步询问是否打开的文件延时修复）：
  * 输入延时根因：交互模式下参数询问用命令行 input()，保存路径用 tkinter
    对话框，而最后一步“是否立即打开生成的 Excel？”又回到裸 input()，形成
    “控制台 ↔ GUI ↔ 控制台”的交互割裂；exe 为 console 模式，控制台窗口在
    后台且通常无焦点/被遮挡，生成完成后用户敲键盘无效、看不到提示，误以为
    程序卡死；
  * 执行延时根因：生成 Excel 阶段在保存对话框关闭后立即同步执行逐周排布与
    大量单元格样式计算（样式对象在循环内逐格新建）+ wb.save 写盘，期间控制台
    无任何进度提示，用户无从判断程序状态；
  * 修复：新增 GUI 优先的“是否打开”确认弹窗（tkinter messagebox.askyesno，
    置顶+立即聚焦+模态事件循环，点击立即返回），彻底消除 CLI/GUI 混用导致
    的输入无效等待；纯命令行/无显示环境（或设置环境变量 DSS_NO_GUI=1）自动
    回退为 input()；生成开始前打印“正在生成”进度提示；渲染循环中的样式对象
    （周次列底色、范围外底色、节点/普通字体、换行对齐）全部提取为复用常量，
    消除逐格新建对象的开销，显著缩短生成→询问前的执行时间；
  * 功能不变：打开/不打开分支与 v3.5 完全一致；GUI 弹窗点击“是/否”立即响应。

v3.7 修订（周期输入方式扩展与打印时间节点拆分）：
  * 周期输入方式扩展：保留原“开始日期+结束日期”方式，新增“开始日期+设计
    计划周期（设计总时间，天）”方式（CLI --period-days / 交互方式2），
    end = start + (periodDays - 1) 天（周期含首尾）；方式2 传入时忽略 --end；
  * 打印时间节点拆分：将 v3.6 合并的“最终节点：出图与图纸归档”区间
    （end-N+1 ~ end）拆分为两个独立节点——“开始出图”（end-N+1）与
    “交付”（end）；交付与整体周期最终日期重合；N=1 时两节点同日不判重合、
    不触发短周期提示；
  * 避开节假日迁移：交付固定于 end，不参与迁移；“开始出图”落休息日且启用
    避开节假日时向前取最近工作日；会审仍安排在 end-N（开始出图前一天）；
  * 版本号统一为 3.7；其余 v3.6 行为（阶段周期规则、阶段结束标注、重合判定、
    样式、表注 JaJa+AI、编制依据）保持不变。

v4.0 修订（GUI 图形交互界面与多风格输出）：
  * 新增 GUI 图形交互界面（tkinter）：默认启动进入 GUI，界面简洁、设计师
    风格、清楚明确；窗口展示软件名称「设计计划时间表生成器」、作者 JaJa+AI、
    版本 v4.0；所有参数均在界面选择或输入：项目名称、设计编号、设计阶段
    （初步设计/施工图设计）、周期输入方式（方式1：开始+结束日期；方式2：
    开始+设计总周期天数）、开始日期、结束日期或总周期天数、出图制作时间
    （默认 1）、避开节假日开关、简洁版开关、表格风格、输出文件保存位置；
    点击「生成时间表」后弹出保存对话框生成 xlsx；GUI 模式严禁混用 input()，
    确认/提醒统一用 tkinter messagebox；
  * 表格多风格：新增 4 种输出风格（桌面与安卓共用）：classic 经典黑白打印
    友好、blue 蓝白商务、green 淡绿清新、warm 暖色简约；风格影响标题/表头/
    阶段行/边框/强调色配色；法定休息日淡绿底（#E2EFDA）、节点重合浅红底
    （#FFC7CE）、斑马纹等既有硬性偏好各风格保留；GUI 提供风格下拉并给出
    简短说明；CLI 新增 --style 参数，缺省 classic（可沿用 config.json 记忆）；
  * 强制 CLI 模式：--cli 参数或环境变量 DSS_NO_GUI=1 时跳过 GUI，回退
    原 CLI 交互模式；完整命令行参数模式（--stage+--start+--end/--period-days）
    行为不变；新增 --version 打印版本号；
  * 版本号统一为 4.0；其余 v3.7 行为（周期输入方式、出图节点拆分、阶段周期
    规则、阶段结束标注、重合判定、样式硬偏好、表注 JaJa+AI、编制依据）不变。

v4.1 修订（专业作业清单）：
  * 新增「是否包含专业作业清单」选项（GUI 复选框 / CLI --work-lists，缺省
    不含，与 --no-work-lists 对应；可沿用 config.json 记忆）；
  * 勾选/指定后，生成的 xlsx 在「设计计划时间表」「节点清单」之外，按每个
    工作专业新增一个工作表（sheet），表名即专业名；sheet 列出该专业各节点
    的日期/阶段/内容与类型（发图/收图），格式参考「节点清单」；
  * 专业范围：初步设计阶段 8 个专业——建筑、总平面、结构、给排水、暖通、
    电气、智能化、经济；施工图阶段 7 个专业——建筑、总平面、结构、给排水、
    暖通、电气、智能化（经济仅初步设计阶段含）；
  * 底色规则：发图（提出）相关行淡橙 #FCE4D6，收图（接收）相关行淡蓝
    #DDEBF7；同一行内容中既有提出又有接收时，按段落拆分到对应专业行，
    并以“发图/收图”类型列明确区分，不混色含糊；
  * 数据来源：优先从《工程设计项目专业配合要求》源文档提取；源文档不可读
    时，从主脚本既有节点数据按专业标签（“XX专业提出/提给XX/接受专业”等）
    分组推导，仅保留文档/节点中真实出现的内容，不新增编造；
  * 桌面与安卓共用同一专业范围与配色；使用说明、帮助文本、版本号、文档
    字符串、GUI 窗口标题全部升 v4.1（版本号 v4.1，作者 JaJa+AI）。

v5.2 修订（覆写节点名称、迁移重合标注与自动打开修复，仅 Windows 桌面版）：
  * 修复：通过「覆写节点」加入的节点，在各阶段计划表格中的节点名称显示
    为节点所在模板文件的文件名（此前误显示模板内节点自身名称）；
  * 修复：因避开节假日迁移导致节点重合时，仅迁移节点自身标注
    “（由X月X日（节假日）迁移而来）”，未迁移的同日常规节点不再被整格
    “由于避开节假日调整导致节点重合”提示覆盖，也不显示任何迁移来源标注；
    迁移重合日期格仍保留浅红标色；
  * 修复：生成后弹窗点击“是”不能自动打开文件的问题——根因为 results
    列表扩展为 6 元组后末尾遍历仍按 5 元组解包导致 ValueError 崩溃；已修正
    解包并保留多重降级打开（os.startfile → cmd start → explorer），
    行为与 v4.2 正常版本一致。

v5.3 修订（修正覆写节点名称显示格式，仅 Windows 桌面版）：
  * 修正 v5.2 bug1 理解偏差：通过「覆写节点」加入的节点，在各阶段计划
    表格中的节点名称显示为「模板文件名：模板内节点名」格式（中间以全角
    冒号“：”分隔），如“幕墙工作（模板文件名）：幕墙施工图提资（模板内
    节点名）”；v5.2 仅显示模板文件名，现已补齐模板内节点名，主表与节点
    清单同步生效。

v5.3.2 修订（13 条优化升级，仅 Windows 桌面版；迁移重合浅红误标修复承自 v5.3.1）：
  * 修复 v5.1 起“避开节假日”重合判定过宽的问题：原判定只要同日至少两
    个节点且其中任一节点带“由X月X日（节假日）迁移而来”标注即整格标浅
    红，导致“阶段开始节点整体从休息日迁移至工作日”（迁至日原本无节点）
    也被误标浅红；
  * 现改为：仅当同日同时存在迁移节点（带迁移标注）与未迁移的原有节点时，
    才判定为“迁移导致与迁至日节点重合”并标浅红；整体迁移（当日全部节点
    都从休息日迁来）不再标浅红。

运行方式（任选其一）：
  1. GUI 图形界面模式（v4.0 默认，推荐）：
        python generate_design_schedule.py
     启动后弹出简洁设计师风格的图形界面，软件名称/作者/版本展示在窗口顶部，
     所有参数在界面选择或输入，点击「生成时间表」弹出保存对话框生成 xlsx；
     设置环境变量 DSS_NO_GUI=1 或加 --cli 参数可强制回退命令行交互模式。
  2. 命令行参数模式：
        python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-02-28
        python generate_design_schedule.py --stage 1 --start 2026-01-05 --period-days 55
        python generate_design_schedule.py --stage 施工图 --start 2026-03-01 --end 2026-05-31 --output D:\计划表.xlsx
        python generate_design_schedule.py --project 测试项目 --code 2026-01 --stage 1 --start 2026-02-10 --end 2026-02-28 --avoid-holidays yes
        python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-01-18 --format brief --drawing-days 1
        python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-01-18 --format both --output D:\plan.xlsx
        python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-01-18 --custom-node "2026-01-12|甲方汇报会|向业主汇报初设方案"
  3. 自定义节假日文件（覆盖内置数据）：
        python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-02-28 --holidays-file my_holidays.json

配置记忆：
  * 配置文件 config.json 与脚本/程序同目录。交互模式启动时读取上次配置，
    各询问项提示默认值，直接回车沿用；命令行未指定参数时也以配置值为默认。
  * 生成完成后自动将本次使用的配置写回 config.json。

软件作者：JaJa+AI 　版本：v4.5中国院版
依赖：Python 3.8+，openpyxl（pip install openpyxl）
=====================================================================
"""

import argparse
import json
import math
import os
import re
import subprocess
import sys
from datetime import date, timedelta

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("缺少依赖 openpyxl，请先执行：pip install openpyxl")
    sys.exit(1)

# 控制台输出编码容错：Windows GBK 控制台无法打印部分字符（如 ⚠）时降级替换，避免崩溃
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(errors="replace")
except Exception:
    pass

# =====================================================================
# 〇、配置记忆（config.json，与脚本/程序同目录）
# =====================================================================

def program_dir():
    """脚本所在目录；打包为 exe 时取 exe 所在目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(program_dir(), "config.json")

CONFIG_SCHEMA_VERSION = 3  # v5.3.3(U2)：config.json schema 版本号

DEFAULT_CONFIG = {
    "schema_version": CONFIG_SCHEMA_VERSION,
    "project_name": "",
    "design_code": "",
    "drawing_days": 1,
    "avoid_holidays": False,
    "format": "full",
    "style": "classic",
    "work_lists": False,
    "languages": [],
    "custom_styles": {},  # v5.3.3(U1)：用户自定义样式（键=样式名，值=样式字典）
    "preflight": {"min_gap_days": 2, "skip_custom": True},  # v5.3.3(U6)：重合预检阈值
}

# =====================================================================
# v5.0：多语种选项（桌面与安卓共用）
# =====================================================================
# 语言代码 -> 中文名称/原语名称（GUI 复选框展示与 CLI --languages 取值）。
# 世界主流语言 12 种，覆盖客户常见需求：英语/日语/韩语/法语/德语/西班牙语/
# 葡萄牙语/俄语/阿拉伯语/意大利语/越南语/泰语。
LANG_OPTIONS = {
    "en": "英语（English）",
    "ja": "日语（日本語）",
    "ko": "韩语（한국어）",
    "fr": "法语（Français）",
    "de": "德语（Deutsch）",
    "es": "西班牙语（Español）",
    "pt": "葡萄牙语（Português）",
    "ru": "俄语（Русский）",
    "ar": "阿拉伯语（العربية）",
    "it": "意大利语（Italiano）",
    "vi": "越南语（Tiếng Việt）",
    "th": "泰语（ไทย）",
}

# 内置中 -> 各语言词典：覆盖标题区、表头、阶段名、常用节点名称、提示语等
# 高频词条。呈现方式：中文 + 所选语言译文（多语依次追加，如 中文　[EN] …　[JA] …）。
# 词典未覆盖的词条保持中文，严禁编造翻译。
# 语言顺序固定：en, ja, ko, fr, de, es, pt, ru, ar, it, vi, th
LANG_DICT = {
    # ---- 标题区 / 表头 / 界面标签 ----
    "设计配合进度计划表": {
        "en": "Design Coordination Schedule", "ja": "設計協調スケジュール", "ko": "설계 조정 일정표",
        "fr": "Calendrier de coordination de conception", "de": "Terminplan der Entwurfskoordination",
        "es": "Cronograma de coordinación de diseño", "pt": "Cronograma de coordenação de projeto",
        "ru": "График координации проектирования", "ar": "جدول تنسيق التصميم",
        "it": "Cronoprogramma di coordinamento progettuale", "vi": "Lịch trình phối hợp thiết kế",
        "th": "ตารางการประสานงานออกแบบ",
    },
    "节点清单汇总表": {
        "en": "Node List Summary", "ja": "ノード一覧表", "ko": "노드 목록 요약표",
        "fr": "Résumé de la liste des nœuds", "de": "Zusammenfassung der Knotenliste",
        "es": "Resumen de lista de nodos", "pt": "Resumo da lista de nós",
        "ru": "Сводный перечень узлов", "ar": "ملخص قائمة العقد",
        "it": "Riepilogo elenco nodi", "vi": "Bảng tổng hợp danh sách mốc", "th": "สรุปรายการจุดสำคัญ",
    },
    "专业作业清单": {
        "en": "Professional Work List", "ja": "専門作業リスト", "ko": "전문 작업 목록",
        "fr": "Liste de travaux professionnels", "de": "Arbeitsliste der Fachdisziplinen",
        "es": "Lista de trabajos profesionales", "pt": "Lista de trabalhos profissionais",
        "ru": "Перечень работ по специальностям", "ar": "قائمة الأعمال المهنية",
        "it": "Elenco lavori professionali", "vi": "Danh sách công việc chuyên môn", "th": "รายการงานวิชาชีพ",
    },
    "项目名称": {
        "en": "Project Name", "ja": "プロジェクト名", "ko": "프로젝트명",
        "fr": "Nom du projet", "de": "Projektname", "es": "Nombre del proyecto",
        "pt": "Nome do projeto", "ru": "Наименование проекта", "ar": "اسم المشروع",
        "it": "Nome del progetto", "vi": "Tên dự án", "th": "ชื่อโครงการ",
    },
    "设计编号": {
        "en": "Design Code", "ja": "設計番号", "ko": "설계 번호",
        "fr": "Code de conception", "de": "Entwurfsnummer", "es": "Código de diseño",
        "pt": "Código do projeto", "ru": "Номер проекта", "ar": "رقم التصميم",
        "it": "Codice di progetto", "vi": "Mã thiết kế", "th": "รหัสการออกแบบ",
    },
    "设计阶段": {
        "en": "Design Stage", "ja": "設計段階", "ko": "설계 단계",
        "fr": "Phase de conception", "de": "Entwurfsphase", "es": "Etapa de diseño",
        "pt": "Fase de projeto", "ru": "Стадия проектирования", "ar": "مرحلة التصميم",
        "it": "Fase di progettazione", "vi": "Giai đoạn thiết kế", "th": "ขั้นตอนการออกแบบ",
    },
    "开始日期": {
        "en": "Start Date", "ja": "開始日", "ko": "시작 날짜",
        "fr": "Date de début", "de": "Startdatum", "es": "Fecha de inicio",
        "pt": "Data de início", "ru": "Дата начала", "ar": "تاريخ البدء",
        "it": "Data di inizio", "vi": "Ngày bắt đầu", "th": "วันที่เริ่มต้น",
    },
    "结束日期": {
        "en": "End Date", "ja": "終了日", "ko": "종료 날짜",
        "fr": "Date de fin", "de": "Enddatum", "es": "Fecha de fin",
        "pt": "Data de término", "ru": "Дата окончания", "ar": "تاريخ الانتهاء",
        "it": "Data di fine", "vi": "Ngày kết thúc", "th": "วันที่สิ้นสุด",
    },
    "设计周期": {
        "en": "Design Period", "ja": "設計周期", "ko": "설계 기간",
        "fr": "Période de conception", "de": "Entwurfszeitraum", "es": "Período de diseño",
        "pt": "Período de projeto", "ru": "Период проектирования", "ar": "فترة التصميم",
        "it": "Periodo di progettazione", "vi": "Chu kỳ thiết kế", "th": "ระยะเวลาการออกแบบ",
    },
    "编制日期": {
        "en": "Date of Preparation", "ja": "作成日", "ko": "작성 날짜",
        "fr": "Date de préparation", "de": "Erstellungsdatum", "es": "Fecha de elaboración",
        "pt": "Data de elaboração", "ru": "Дата составления", "ar": "تاريخ الإعداد",
        "it": "Data di redazione", "vi": "Ngày lập", "th": "วันที่จัดทำ",
    },
    "生成日期": {
        "en": "Generated On", "ja": "生成日", "ko": "생성 날짜",
        "fr": "Date de génération", "de": "Erzeugungsdatum", "es": "Fecha de generación",
        "pt": "Data de geração", "ru": "Дата генерации", "ar": "تاريخ الإنشاء",
        "it": "Data di generazione", "vi": "Ngày tạo", "th": "วันที่สร้าง",
    },
    "节假日处理": {
        "en": "Holiday Handling", "ja": "祝日対応", "ko": "휴일 처리",
        "fr": "Gestion des jours fériés", "de": "Feiertagsbehandlung", "es": "Gestión de festivos",
        "pt": "Tratamento de feriados", "ru": "Обработка праздников", "ar": "معالجة العطلات",
        "it": "Gestione festivi", "vi": "Xử lý ngày lễ", "th": "การจัดการวันหยุด",
    },
    "输出版本": {
        "en": "Output Version", "ja": "出力バージョン", "ko": "출력 버전",
        "fr": "Version de sortie", "de": "Ausgabeversion", "es": "Versión de salida",
        "pt": "Versão de saída", "ru": "Версия вывода", "ar": "إصدار المخرجات",
        "it": "Versione di output", "vi": "Phiên bản xuất", "th": "เวอร์ชันผลลัพธ์",
    },
    "完整版": {
        "en": "Full Version", "ja": "完全版", "ko": "전체 버전",
        "fr": "Version complète", "de": "Vollversion", "es": "Versión completa",
        "pt": "Versão completa", "ru": "Полная версия", "ar": "النسخة الكاملة",
        "it": "Versione completa", "vi": "Bản đầy đủ", "th": "เวอร์ชันเต็ม",
    },
    "简洁版": {
        "en": "Brief Version", "ja": "簡潔版", "ko": "간결 버전",
        "fr": "Version concise", "de": "Kurzversion", "es": "Versión resumida",
        "pt": "Versão resumida", "ru": "Краткая версия", "ar": "النسخة المختصرة",
        "it": "Versione sintetica", "vi": "Bản tóm tắt", "th": "เวอร์ชันย่อ",
    },
    "周次": {
        "en": "Week", "ja": "週次", "ko": "주차",
        "fr": "Semaine", "de": "Woche", "es": "Semana",
        "pt": "Semana", "ru": "Неделя", "ar": "الأسبوع",
        "it": "Settimana", "vi": "Tuần", "th": "สัปดาห์",
    },
    "周一": {
        "en": "Mon", "ja": "月曜", "ko": "월요일", "fr": "Lun", "de": "Mo",
        "es": "Lun", "pt": "Seg", "ru": "Пн", "ar": "الاثنين", "it": "Lun",
        "vi": "T2", "th": "จ.",
    },
    "周二": {
        "en": "Tue", "ja": "火曜", "ko": "화요일", "fr": "Mar", "de": "Di",
        "es": "Mar", "pt": "Ter", "ru": "Вт", "ar": "الثلاثاء", "it": "Mar",
        "vi": "T3", "th": "อ.",
    },
    "周三": {
        "en": "Wed", "ja": "水曜", "ko": "수요일", "fr": "Mer", "de": "Mi",
        "es": "Mié", "pt": "Qua", "ru": "Ср", "ar": "الأربعاء", "it": "Mer",
        "vi": "T4", "th": "พ.",
    },
    "周四": {
        "en": "Thu", "ja": "木曜", "ko": "목요일", "fr": "Jeu", "de": "Do",
        "es": "Jue", "pt": "Qui", "ru": "Чт", "ar": "الخميس", "it": "Gio",
        "vi": "T5", "th": "พฤ.",
    },
    "周五": {
        "en": "Fri", "ja": "金曜", "ko": "금요일", "fr": "Ven", "de": "Fr",
        "es": "Vie", "pt": "Sex", "ru": "Пт", "ar": "الجمعة", "it": "Ven",
        "vi": "T6", "th": "ศ.",
    },
    "周六": {
        "en": "Sat", "ja": "土曜", "ko": "토요일", "fr": "Sam", "de": "Sa",
        "es": "Sáb", "pt": "Sáb", "ru": "Сб", "ar": "السبت", "it": "Sab",
        "vi": "T7", "th": "ส.",
    },
    "周日": {
        "en": "Sun", "ja": "日曜", "ko": "일요일", "fr": "Dim", "de": "So",
        "es": "Dom", "pt": "Dom", "ru": "Вс", "ar": "الأحد", "it": "Dom",
        "vi": "CN", "th": "อา.",
    },
    "序号": {
        "en": "No.", "ja": "番号", "ko": "번호",
        "fr": "N°", "de": "Nr.", "es": "N.º", "pt": "N.º", "ru": "№",
        "ar": "الرقم", "it": "N.", "vi": "STT", "th": "ลำดับ",
    },
    "日期": {
        "en": "Date", "ja": "日付", "ko": "날짜",
        "fr": "Date", "de": "Datum", "es": "Fecha", "pt": "Data", "ru": "Дата",
        "ar": "التاريخ", "it": "Data", "vi": "Ngày", "th": "วันที่",
    },
    "星期": {
        "en": "Weekday", "ja": "曜日", "ko": "요일",
        "fr": "Jour de semaine", "de": "Wochentag", "es": "Día de la semana",
        "pt": "Dia da semana", "ru": "День недели", "ar": "اليوم",
        "it": "Giorno della settimana", "vi": "Thứ", "th": "วัน",
    },
    "所属阶段": {
        "en": "Stage", "ja": "所属段階", "ko": "소속 단계",
        "fr": "Phase", "de": "Phase", "es": "Etapa", "pt": "Fase", "ru": "Стадия",
        "ar": "المرحلة", "it": "Fase", "vi": "Giai đoạn", "th": "ขั้นตอน",
    },
    "节点名称": {
        "en": "Node Name", "ja": "ノード名", "ko": "노드 이름",
        "fr": "Nom du nœud", "de": "Knotenname", "es": "Nombre del nodo",
        "pt": "Nome do nó", "ru": "Название узла", "ar": "اسم العقدة",
        "it": "Nome del nodo", "vi": "Tên mốc", "th": "ชื่อจุด",
    },
    "具体内容或章节编号": {
        "en": "Details / Section No.", "ja": "具体内容／章番号", "ko": "구체 내용/절 번호",
        "fr": "Détails / N° de section", "de": "Details / Abschnitt-Nr.", "es": "Detalles / N.º de sección",
        "pt": "Detalhes / N.º de seção", "ru": "Содержание / № раздела", "ar": "التفاصيل / رقم القسم",
        "it": "Dettagli / N. sezione", "vi": "Chi tiết / Số mục", "th": "รายละเอียด / หมายเลขข้อ",
    },
    "里程碑标记": {
        "en": "Milestone Mark", "ja": "マイルストーン", "ko": "이정표 표시",
        "fr": "Jalon", "de": "Meilenstein", "es": "Hito",
        "pt": "Marco", "ru": "Контрольная точка", "ar": "علامة مرجعية",
        "it": "Traguardo", "vi": "Mốc quan trọng", "th": "จุดสำคัญ",
    },
    "里程碑": {
        "en": "Milestone", "ja": "マイルストーン", "ko": "이정표",
        "fr": "Jalon", "de": "Meilenstein", "es": "Hito", "pt": "Marco",
        "ru": "Контрольная точка", "ar": "مرجعية", "it": "Traguardo",
        "vi": "Mốc quan trọng", "th": "จุดสำคัญ",
    },
    "内容/事项": {
        "en": "Content / Item", "ja": "内容／事項", "ko": "내용/사항",
        "fr": "Contenu / Élément", "de": "Inhalt / Position", "es": "Contenido / Elemento",
        "pt": "Conteúdo / Item", "ru": "Содержание / Позиция", "ar": "المحتوى / البند",
        "it": "Contenuto / Voce", "vi": "Nội dung / Hạng mục", "th": "เนื้อหา / รายการ",
    },
    "类型": {
        "en": "Type", "ja": "種類", "ko": "유형",
        "fr": "Type", "de": "Typ", "es": "Tipo", "pt": "Tipo", "ru": "Тип",
        "ar": "النوع", "it": "Tipo", "vi": "Loại", "th": "ประเภท",
    },
    "发图": {
        "en": "Issue Drawings", "ja": "図面提出", "ko": "도면 발행",
        "fr": "Émission des plans", "de": "Planausgabe", "es": "Emisión de planos",
        "pt": "Emissão de desenhos", "ru": "Выдача чертежей", "ar": "إصدار المخططات",
        "it": "Emissione disegni", "vi": "Phát bản vẽ", "th": "ส่งแบบ",
    },
    "收图": {
        "en": "Receive Drawings", "ja": "図面受領", "ko": "도면 수령",
        "fr": "Réception des plans", "de": "Planempfang", "es": "Recepción de planos",
        "pt": "Recebimento de desenhos", "ru": "Получение чертежей", "ar": "استلام المخططات",
        "it": "Ricezione disegni", "vi": "Nhận bản vẽ", "th": "รับแบบ",
    },
    "合计": {
        "en": "Total", "ja": "合計", "ko": "합계",
        "fr": "Total", "de": "Gesamt", "es": "Total", "pt": "Total", "ru": "Итого",
        "ar": "الإجمالي", "it": "Totale", "vi": "Tổng cộng", "th": "รวม",
    },
    "阶段结束": {
        "en": "Stage End", "ja": "段階終了", "ko": "단계 종료",
        "fr": "Fin de phase", "de": "Phasenende", "es": "Fin de etapa",
        "pt": "Fim de fase", "ru": "Окончание стадии", "ar": "نهاية المرحلة",
        "it": "Fine fase", "vi": "Kết thúc giai đoạn", "th": "สิ้นสุดขั้นตอน",
    },
    "休息日": {
        "en": "Rest Day", "ja": "休日", "ko": "휴일",
        "fr": "Jour de repos", "de": "Ruhetag", "es": "Día de descanso",
        "pt": "Dia de descanso", "ru": "Выходной", "ar": "يوم راحة",
        "it": "Giorno di riposo", "vi": "Ngày nghỉ", "th": "วันหยุด",
    },
    "工作日": {
        "en": "Working Day", "ja": "営業日", "ko": "근무일",
        "fr": "Jour ouvrable", "de": "Werktag", "es": "Día laborable",
        "pt": "Dia útil", "ru": "Рабочий день", "ar": "يوم عمل",
        "it": "Giorno lavorativo", "vi": "Ngày làm việc", "th": "วันทำงาน",
    },
    "避开节假日": {
        "en": "Avoid Holidays", "ja": "祝日を避ける", "ko": "휴일 회피",
        "fr": "Éviter les jours fériés", "de": "Feiertage vermeiden", "es": "Evitar festivos",
        "pt": "Evitar feriados", "ru": "Избегать праздников", "ar": "تجنب العطلات",
        "it": "Evita festivi", "vi": "Tránh ngày lễ", "th": "หลีกเลี่ยงวันหยุด",
    },
    "节点迁移至最近工作日": {
        "en": "Node moved to nearest working day", "ja": "ノードを直近の営業日に移動",
        "ko": "노드를 가장 가까운 근무일로 이동",
        "fr": "Nœud déplacé au jour ouvrable le plus proche",
        "de": "Knoten auf nächsten Werktag verschoben",
        "es": "Nodo movido al día laborable más cercano",
        "pt": "Nó movido para o dia útil mais próximo",
        "ru": "Узел перенесён на ближайший рабочий день",
        "ar": "تم نقل العقدة إلى أقرب يوم عمل",
        "it": "Nodo spostato al giorno lavorativo più vicino",
        "vi": "Đã chuyển mốc đến ngày làm việc gần nhất",
        "th": "ย้ายจุดไปยังวันทำงานที่ใกล้ที่สุด",
    },
    "初步设计阶段": {
        "en": "Preliminary Design Stage", "ja": "基本設計段階", "ko": "기본 설계 단계",
        "fr": "Phase d'avant-projet", "de": "Vorprojektphase", "es": "Etapa de diseño preliminar",
        "pt": "Fase de projeto preliminar", "ru": "Стадия эскизного проектирования",
        "ar": "مرحلة التصميم الأولي", "it": "Fase di progettazione preliminare",
        "vi": "Giai đoạn thiết kế cơ sở", "th": "ขั้นตอนการออกแบบเบื้องต้น",
    },
    "施工图设计阶段": {
        "en": "Construction Drawing Stage", "ja": "実施設計段階", "ko": "실시 설계 단계",
        "fr": "Phase de conception détaillée", "de": "Ausführungsplanungsphase",
        "es": "Etapa de diseño de ejecución", "pt": "Fase de projeto executivo",
        "ru": "Стадия рабочей документации", "ar": "مرحلة المخططات التنفيذية",
        "it": "Fase di progettazione esecutiva", "vi": "Giai đoạn thiết kế thi công",
        "th": "ขั้นตอนการเขียนแบบก่อสร้าง",
    },
    "项目概况": {
        "en": "Project Overview", "ja": "プロジェクト概要", "ko": "프로젝트 개요",
        "fr": "Aperçu du projet", "de": "Projektübersicht", "es": "Resumen del proyecto",
        "pt": "Visão geral do projeto", "ru": "Обзор проекта", "ar": "نظرة عامة على المشروع",
        "it": "Panoramica del progetto", "vi": "Tổng quan dự án", "th": "ภาพรวมโครงการ",
    },
    "节点说明": {
        "en": "Node Notes", "ja": "ノード説明", "ko": "노드 설명",
        "fr": "Notes sur les nœuds", "de": "Knotenhinweise", "es": "Notas sobre nodos",
        "pt": "Notas sobre nós", "ru": "Пояснения к узлам", "ar": "ملاحظات العقد",
        "it": "Note sui nodi", "vi": "Ghi chú mốc", "th": "หมายเหตุจุด",
    },
    "阶段时间周期": {
        "en": "Stage Time Periods", "ja": "段階期間", "ko": "단계 기간",
        "fr": "Périodes de phase", "de": "Phasenzeiträume", "es": "Períodos de etapa",
        "pt": "Períodos de fase", "ru": "Периоды стадий", "ar": "فترات المراحل",
        "it": "Periodi di fase", "vi": "Chu kỳ giai đoạn", "th": "ระยะเวลาขั้นตอน",
    },
    "阶段周期表达规则": {
        "en": "Stage Period Rules", "ja": "段階周期表現ルール", "ko": "단계 기간 표현 규칙",
        "fr": "Règles de période de phase", "de": "Regeln für Phasenzeiträume",
        "es": "Reglas de períodos de etapa", "pt": "Regras de períodos de fase",
        "ru": "Правила выражения периодов", "ar": "قواعد فترات المراحل",
        "it": "Regole periodi di fase", "vi": "Quy tắc chu kỳ giai đoạn", "th": "กฎระยะเวลาขั้นตอน",
    },
    "图例": {
        "en": "Legend", "ja": "凡例", "ko": "범례",
        "fr": "Légende", "de": "Legende", "es": "Leyenda", "pt": "Legenda",
        "ru": "Условные обозначения", "ar": "دليل الرموز", "it": "Legenda",
        "vi": "Chú giải", "th": "คำอธิบายสัญลักษณ์",
    },
    "编制依据": {
        "en": "Basis of Preparation", "ja": "作成根拠", "ko": "작성 근거",
        "fr": "Base de préparation", "de": "Erstellungsgrundlage", "es": "Base de elaboración",
        "pt": "Base de elaboração", "ru": "Основание составления", "ar": "أساس الإعداد",
        "it": "Base di redazione", "vi": "Căn cứ lập", "th": "หลักฐานการจัดทำ",
    },
    "软件信息": {
        "en": "Software Info", "ja": "ソフト情報", "ko": "소프트웨어 정보",
        "fr": "Info logiciel", "de": "Software-Info", "es": "Información del software",
        "pt": "Informações do software", "ru": "О программе", "ar": "معلومات البرنامج",
        "it": "Info software", "vi": "Thông tin phần mềm", "th": "ข้อมูลซอฟต์แวร์",
    },
    "阶段结束标注": {
        "en": "Stage End Mark", "ja": "段階終了マーク", "ko": "단계 종료 표시",
        "fr": "Marque de fin de phase", "de": "Phasenende-Markierung", "es": "Marca de fin de etapa",
        "pt": "Marca de fim de fase", "ru": "Метка окончания стадии", "ar": "علامة نهاية المرحلة",
        "it": "Marcatura fine fase", "vi": "Ký hiệu kết thúc giai đoạn", "th": "เครื่องหมายสิ้นสุดขั้นตอน",
    },
    "（休息日）": {
        "en": "(Rest day)", "ja": "（休日）", "ko": "（휴일）",
        "fr": "(Jour de repos)", "de": "(Ruhetag)", "es": "(Día de descanso)",
        "pt": "(Dia de descanso)", "ru": "(Выходной)", "ar": "(يوم راحة)",
        "it": "(Giorno di riposo)", "vi": "(Ngày nghỉ)", "th": "(วันหยุด)",
    },
    "（工作日）": {
        "en": "(Working day)", "ja": "（営業日）", "ko": "（근무일）",
        "fr": "(Jour ouvrable)", "de": "(Werktag)", "es": "(Día laborable)",
        "pt": "(Dia útil)", "ru": "(Рабочий день)", "ar": "(يوم عمل)",
        "it": "(Giorno lavorativo)", "vi": "(Ngày làm việc)", "th": "(วันทำงาน)",
    },
    # ---- 设计阶段名（初步设计 6 阶段 / 施工图 8 阶段）----
    "第一阶段 初设前期阶段": {
        "en": "Stage 1: Pre-design Period", "ja": "第1段階 基本設計前期", "ko": "1단계 초기 설계 단계",
        "fr": "Phase 1 : Période d'avant-projet", "de": "Phase 1: Vorprojektphase",
        "es": "Etapa 1: Período pre-diseño", "pt": "Fase 1: Período pré-projeto",
        "ru": "Стадия 1: Предпроектный период", "ar": "المرحلة 1: فترة ما قبل التصميم",
        "it": "Fase 1: Periodo pre-progettuale", "vi": "Giai đoạn 1: Trước thiết kế cơ sở",
        "th": "ขั้นตอนที่ 1: ช่วงก่อนออกแบบ",
    },
    "第二阶段 第一版作业图阶段": {
        "en": "Stage 2: First Draft Working Drawings", "ja": "第2段階 第一版作業図", "ko": "2단계 1차 작업도 단계",
        "fr": "Phase 2 : Première version des plans de travail", "de": "Phase 2: Erste Arbeitsplanfassung",
        "es": "Etapa 2: Primera versión de planos de trabajo", "pt": "Fase 2: Primeira versão dos desenhos de trabalho",
        "ru": "Стадия 2: Первая версия рабочих чертежей", "ar": "المرحلة 2: النسخة الأولى من مخططات العمل",
        "it": "Fase 2: Prima versione disegni di lavoro", "vi": "Giai đoạn 2: Bản vẽ làm việc lần 1",
        "th": "ขั้นตอนที่ 2: แบบงานครั้งที่ 1",
    },
    "第三阶段 第二版作业图阶段": {
        "en": "Stage 3: Second Draft Working Drawings", "ja": "第3段階 第二版作業図", "ko": "3단계 2차 작업도 단계",
        "fr": "Phase 3 : Deuxième version des plans de travail", "de": "Phase 3: Zweite Arbeitsplanfassung",
        "es": "Etapa 3: Segunda versión de planos de trabajo", "pt": "Fase 3: Segunda versão dos desenhos de trabalho",
        "ru": "Стадия 3: Вторая версия рабочих чертежей", "ar": "المرحلة 3: النسخة الثانية من مخططات العمل",
        "it": "Fase 3: Seconda versione disegni di lavoro", "vi": "Giai đoạn 3: Bản vẽ làm việc lần 2",
        "th": "ขั้นตอนที่ 3: แบบงานครั้งที่ 2",
    },
    "第四阶段 管道综合与对图阶段": {
        "en": "Stage 4: Pipe Coordination & Alignment", "ja": "第4段階 配管統合・照合", "ko": "4단계 배관 종합·대조 단계",
        "fr": "Phase 4 : Synthèse et vérification des réseaux", "de": "Phase 4: Rohrleitungsintegration und Abgleich",
        "es": "Etapa 4: Integración y verificación de redes", "pt": "Fase 4: Integração e verificação de redes",
        "ru": "Стадия 4: Комплексная увязка трубопроводов", "ar": "المرحلة 4: تجميع ومطابقة الأنابيب",
        "it": "Fase 4: Integrazione e verifica impianti", "vi": "Giai đoạn 4: Tổng hợp đường ống và đối chiếu",
        "th": "ขั้นตอนที่ 4: รวมระบบท่อและตรวจเทียบ",
    },
    "第五阶段 会审与会签": {
        "en": "Stage 5: Joint Review & Sign-off", "ja": "第5段階 合同審査・署名", "ko": "5단계 합동 심의·서명",
        "fr": "Phase 5 : Revue conjointe et signature", "de": "Phase 5: Gemeinsame Prüfung und Freigabe",
        "es": "Etapa 5: Revisión conjunta y firma", "pt": "Fase 5: Revisão conjunta e assinatura",
        "ru": "Стадия 5: Совместная проверка и подписание", "ar": "المرحلة 5: المراجعة المشتركة والتوقيع",
        "it": "Fase 5: Revisione congiunta e firma", "vi": "Giai đoạn 5: Hội thẩm và ký duyệt",
        "th": "ขั้นตอนที่ 5: ประชุมร่วมและลงนาม",
    },
    "第六阶段 图纸归档": {
        "en": "Stage 6: Drawing Archiving", "ja": "第6段階 図面保管", "ko": "6단계 도면 보관",
        "fr": "Phase 6 : Archivage des plans", "de": "Phase 6: Planarchivierung",
        "es": "Etapa 6: Archivo de planos", "pt": "Fase 6: Arquivamento de desenhos",
        "ru": "Стадия 6: Архивация чертежей", "ar": "المرحلة 6: أرشفة المخططات",
        "it": "Fase 6: Archiviazione disegni", "vi": "Giai đoạn 6: Lưu trữ bản vẽ",
        "th": "ขั้นตอนที่ 6: เก็บแบบ",
    },
    "第一阶段 施工图前期阶段": {
        "en": "Stage 1: Pre-construction Drawing Period", "ja": "第1段階 実施設計前期", "ko": "1단계 실시 설계 전기 단계",
        "fr": "Phase 1 : Période préparatoire", "de": "Phase 1: Vorbereitungsphase",
        "es": "Etapa 1: Período preparatorio", "pt": "Fase 1: Período preparatório",
        "ru": "Стадия 1: Подготовительный период", "ar": "المرحلة 1: الفترة التحضيرية",
        "it": "Fase 1: Periodo preparatorio", "vi": "Giai đoạn 1: Trước khi thiết kế thi công",
        "th": "ขั้นตอนที่ 1: ช่วงก่อนเขียนแบบก่อสร้าง",
    },
    "第五阶段 提清资料阶段": {
        "en": "Stage 5: Final Data Submission", "ja": "第5段階 最終資料提出", "ko": "5단계 최종 자료 제출 단계",
        "fr": "Phase 5 : Remise finale des données", "de": "Phase 5: Endgültige Datenabgabe",
        "es": "Etapa 5: Entrega final de datos", "pt": "Fase 5: Entrega final de dados",
        "ru": "Стадия 5: Окончательная передача данных", "ar": "المرحلة 5: تسليم البيانات النهائية",
        "it": "Fase 5: Consegna finale dati", "vi": "Giai đoạn 5: Trình dữ liệu cuối cùng",
        "th": "ขั้นตอนที่ 5: ส่งข้อมูลขั้นสุดท้าย",
    },
    "第六阶段 建筑作业底图阶段": {
        "en": "Stage 6: Base Drawing Stage", "ja": "第6段階 建築ベース図", "ko": "6단계 건축 작업 베이스 단계",
        "fr": "Phase 6 : Plan de base architectural", "de": "Phase 6: Architektur-Basisdaten",
        "es": "Etapa 6: Plano base de arquitectura", "pt": "Fase 6: Desenho base de arquitetura",
        "ru": "Стадия 6: Архитектурная основа", "ar": "المرحلة 6: المخطط الأساسي المعماري",
        "it": "Fase 6: Base architettonica", "vi": "Giai đoạn 6: Bản vẽ nền kiến trúc",
        "th": "ขั้นตอนที่ 6: แบบพื้นฐานสถาปัตยกรรม",
    },
    "第七阶段 会审与会签": {
        "en": "Stage 7: Joint Review & Sign-off", "ja": "第7段階 合同審査・署名", "ko": "7단계 합동 심의·서명",
        "fr": "Phase 7 : Revue conjointe et signature", "de": "Phase 7: Gemeinsame Prüfung und Freigabe",
        "es": "Etapa 7: Revisión conjunta y firma", "pt": "Fase 7: Revisão conjunta e assinatura",
        "ru": "Стадия 7: Совместная проверка и подписание", "ar": "المرحلة 7: المراجعة المشتركة والتوقيع",
        "it": "Fase 7: Revisione congiunta e firma", "vi": "Giai đoạn 7: Hội thẩm và ký duyệt",
        "th": "ขั้นตอนที่ 7: ประชุมร่วมและลงนาม",
    },
    "第八阶段 图纸归档": {
        "en": "Stage 8: Drawing Archiving", "ja": "第8段階 図面保管", "ko": "8단계 도면 보관",
        "fr": "Phase 8 : Archivage des plans", "de": "Phase 8: Planarchivierung",
        "es": "Etapa 8: Archivo de planos", "pt": "Fase 8: Arquivamento de desenhos",
        "ru": "Стадия 8: Архивация чертежей", "ar": "المرحلة 8: أرشفة المخططات",
        "it": "Fase 8: Archiviazione disegni", "vi": "Giai đoạn 8: Lưu trữ bản vẽ",
        "th": "ขั้นตอนที่ 8: เก็บแบบ",
    },
    # ---- 常用节点名称 ----
    "首次专业协调会": {
        "en": "1st Professional Coordination Meeting", "ja": "第1回専門調整会議", "ko": "제1차 전문 조정 회의",
        "fr": "1re réunion de coordination", "de": "1. Fachkoordinationssitzung",
        "es": "1.ª reunión de coordinación", "pt": "1.ª reunião de coordenação",
        "ru": "Первое совещание по координации", "ar": "الاجتماع الأول للتنسيق المهني",
        "it": "1ª riunione di coordinamento", "vi": "Họp phối hợp chuyên môn lần 1",
        "th": "ประชุมประสานงานวิชาชีพครั้งที่ 1",
    },
    "第二次专业协调会": {
        "en": "2nd Professional Coordination Meeting", "ja": "第2回専門調整会議", "ko": "제2차 전문 조정 회의",
        "fr": "2e réunion de coordination", "de": "2. Fachkoordinationssitzung",
        "es": "2.ª reunión de coordinación", "pt": "2.ª reunião de coordenação",
        "ru": "Второе совещание по координации", "ar": "الاجتماع الثاني للتنسيق المهني",
        "it": "2ª riunione di coordinamento", "vi": "Họp phối hợp chuyên môn lần 2",
        "th": "ประชุมประสานงานวิชาชีพครั้งที่ 2",
    },
    "第三次专业协调会": {
        "en": "3rd Professional Coordination Meeting", "ja": "第3回専門調整会議", "ko": "제3차 전문 조정 회의",
        "fr": "3e réunion de coordination", "de": "3. Fachkoordinationssitzung",
        "es": "3.ª reunión de coordinación", "pt": "3.ª reunião de coordenação",
        "ru": "Третье совещание по координации", "ar": "الاجتماع الثالث للتنسيق المهني",
        "it": "3ª riunione di coordinamento", "vi": "Họp phối hợp chuyên môn lần 3",
        "th": "ประชุมประสานงานวิชาชีพครั้งที่ 3",
    },
    "建筑方案图": {
        "en": "Architectural Scheme Drawings", "ja": "建築基本計画図", "ko": "건축 기본안 도면",
        "fr": "Plans du projet architectural", "de": "Architektonische Entwurfspläne",
        "es": "Planos de esquema arquitectónico", "pt": "Plantas do esquema arquitetônico",
        "ru": "Архитектурные схемные чертежи", "ar": "مخططات التصميم المعماري",
        "it": "Elaborati di schema architettonico", "vi": "Bản vẽ phương án kiến trúc",
        "th": "แบบสถาปัตยกรรมขั้นต้น",
    },
    "总图方案图": {
        "en": "Site Master Scheme Drawings", "ja": "総合計画スキーム図", "ko": "대지 종합 안 도면",
        "fr": "Plans du schéma d'ensemble", "de": "Geländeplan-Schemata",
        "es": "Planos de esquema del conjunto", "pt": "Plantas do esquema do conjunto",
        "ru": "Генплан схемы", "ar": "مخططات الموقع العام",
        "it": "Elaborati di schema generale", "vi": "Bản vẽ phương án tổng mặt bằng",
        "th": "แบบผังรวม",
    },
    "结构专业反馈意见": {
        "en": "Structural Feedback", "ja": "構造専門のフィードバック", "ko": "구조 전문 피드백",
        "fr": "Retour de la structure", "de": "Rückmeldung der Tragwerksplanung",
        "es": "Retroalimentación de estructura", "pt": "Retorno da estrutura",
        "ru": "Замечания конструкторов", "ar": "ملاحظات التخصص الإنشائي",
        "it": "Feedback della struttura", "vi": "Ý kiến phản hồi kết cấu",
        "th": "ข้อเสนอแนะฝ่ายโครงสร้าง",
    },
    "机电专业反馈意见": {
        "en": "MEP Feedback", "ja": "設備専門のフィードバック", "ko": "설비 전문 피드백",
        "fr": "Retour des équipements", "de": "Rückmeldung der Technischen Gebäudeausrüstung",
        "es": "Retroalimentación de instalaciones", "pt": "Retorno das instalações",
        "ru": "Замечания инженерных систем", "ar": "ملاحظات التخصصات الكهربائية والميكانيكية",
        "it": "Feedback impianti", "vi": "Ý kiến phản hồi hệ thống kỹ thuật",
        "th": "ข้อเสนอแนะฝ่ายระบบวิศวกรรม",
    },
    "建筑作业图（第一版）": {
        "en": "Working Drawings (1st Version)", "ja": "作業図（第一版）", "ko": "작업도 (1차)",
        "fr": "Plans de travail (1re version)", "de": "Arbeitspläne (1. Fassung)",
        "es": "Planos de trabajo (1.ª versión)", "pt": "Desenhos de trabalho (1.ª versão)",
        "ru": "Рабочие чертежи (первая версия)", "ar": "مخططات العمل (النسخة الأولى)",
        "it": "Disegni di lavoro (1ª versione)", "vi": "Bản vẽ làm việc (lần 1)",
        "th": "แบบงาน (ครั้งที่ 1)",
    },
    "建筑第二版作业图": {
        "en": "2nd Version Working Drawings", "ja": "第二版作業図", "ko": "2차 작업도",
        "fr": "2e version des plans de travail", "de": "2. Fassung der Arbeitspläne",
        "es": "2.ª versión de planos de trabajo", "pt": "2.ª versão dos desenhos de trabalho",
        "ru": "Рабочие чертежи (вторая версия)", "ar": "النسخة الثانية من مخططات العمل",
        "it": "2ª versione disegni di lavoro", "vi": "Bản vẽ làm việc lần 2",
        "th": "แบบงานครั้งที่ 2",
    },
    "总图作业图": {
        "en": "Site Working Drawings", "ja": "総合計画作業図", "ko": "대지 작업도",
        "fr": "Plans de travail du site", "de": "Gelände-Arbeitspläne",
        "es": "Planos de trabajo del conjunto", "pt": "Desenhos de trabalho do conjunto",
        "ru": "Рабочие чертежи генплана", "ar": "مخططات العمل للموقع العام",
        "it": "Disegni di lavoro del sito", "vi": "Bản vẽ làm việc tổng mặt bằng",
        "th": "แบบงานผังรวม",
    },
    "机电专业确认与提清资料": {
        "en": "MEP Confirmation & Final Data", "ja": "設備確認・最終資料", "ko": "설비 확인 및 최종 자료",
        "fr": "Confirmation et données finales des équipements", "de": "Bestätigung und Abschlussdaten TGA",
        "es": "Confirmación y datos finales de instalaciones", "pt": "Confirmação e dados finais das instalações",
        "ru": "Подтверждение и окончательные данные инженерных систем", "ar": "تأكيد وتسليم البيانات النهائية للأنظمة",
        "it": "Conferma e dati finali impianti", "vi": "Xác nhận và trình dữ liệu cuối hệ thống",
        "th": "ยืนยันและส่งข้อมูลสุดท้ายระบบวิศวกรรม",
    },
    "建筑资料的局部变更": {
        "en": "Partial Changes to Architectural Data", "ja": "建築資料の部分変更", "ko": "건축 자료 부분 변경",
        "fr": "Modifications partielles des données architecturales", "de": "Teilweise Änderungen der Architekturdaten",
        "es": "Cambios parciales en datos de arquitectura", "pt": "Alterações parciais nos dados de arquitetura",
        "ru": "Частичные изменения архитектурных данных", "ar": "تعديلات جزئية على البيانات المعمارية",
        "it": "Modifiche parziali ai dati architettonici", "vi": "Thay đổi cục bộ dữ liệu kiến trúc",
        "th": "แก้ไขข้อมูลสถาปัตยกรรมบางส่วน",
    },
    "管道综合与对图": {
        "en": "Pipe Coordination & Alignment", "ja": "配管統合・照合", "ko": "배관 종합·대조",
        "fr": "Synthèse et vérification des réseaux", "de": "Rohrleitungsintegration und Abgleich",
        "es": "Integración y verificación de redes", "pt": "Integração e verificação de redes",
        "ru": "Комплексная увязка трубопроводов", "ar": "تجميع ومطابقة الأنابيب",
        "it": "Integrazione e verifica impianti", "vi": "Tổng hợp đường ống và đối chiếu",
        "th": "รวมระบบท่อและตรวจเทียบ",
    },
    "机电专业对总图所提资料的确认与局部变更": {
        "en": "MEP Confirmation & Changes to Site Data", "ja": "設備の総図資料確認と局部変更", "ko": "설비의 대지 자료 확인 및 부분 변경",
        "fr": "Confirmation et modifications des données de site", "de": "Bestätigung und Änderungen der Geländedaten",
        "es": "Confirmación y cambios en datos del conjunto", "pt": "Confirmação e alterações dos dados do conjunto",
        "ru": "Подтверждение и изменения данных генплана", "ar": "تأكيد وتعديل بيانات الموقع العام",
        "it": "Conferma e modifiche dati sito", "vi": "Xác nhận và điều chỉnh dữ liệu tổng mặt bằng",
        "th": "ยืนยันและแก้ไขข้อมูลผังรวม",
    },
    "专业间核对、会审与会签": {
        "en": "Inter-disciplinary Check, Review & Sign-off", "ja": "専門間照合・合同審査・署名", "ko": "전문 간 대조·합동 심의·서명",
        "fr": "Vérification interdisciplinaire, revue et signature", "de": "Interdisziplinäre Prüfung, Review und Freigabe",
        "es": "Verificación interdisciplinar, revisión y firma", "pt": "Verificação interdisciplinar, revisão e assinatura",
        "ru": "Междисциплинарная проверка, совместная приёмка и подписание", "ar": "المراجعة المشتركة والتوقيع بين التخصصات",
        "it": "Verifica interdisciplinare, revisione e firma", "vi": "Đối chiếu liên chuyên môn, hội thẩm và ký duyệt",
        "th": "ตรวจสอบข้ามวิชาชีพ ประชุมร่วมและลงนาม",
    },
    "图纸归档": {
        "en": "Drawing Archiving", "ja": "図面保管", "ko": "도면 보관",
        "fr": "Archivage des plans", "de": "Planarchivierung", "es": "Archivo de planos",
        "pt": "Arquivamento de desenhos", "ru": "Архивация чертежей", "ar": "أرشفة المخططات",
        "it": "Archiviazione disegni", "vi": "Lưu trữ bản vẽ", "th": "เก็บแบบ",
    },
    "建筑初步设计图": {
        "en": "Preliminary Architectural Drawings", "ja": "建築基本設計図", "ko": "건축 기본 설계도",
        "fr": "Plans d'avant-projet architectural", "de": "Architektonische Vorprojektpläne",
        "es": "Planos preliminares de arquitectura", "pt": "Desenhos preliminares de arquitetura",
        "ru": "Архитектурные эскизные чертежи", "ar": "المخططات المعمارية الأولية",
        "it": "Elaborati preliminari architettonici", "vi": "Bản vẽ thiết kế cơ sở kiến trúc",
        "th": "แบบสถาปัตยกรรมเบื้องต้น",
    },
    "总图初步设计图": {
        "en": "Preliminary Site Drawings", "ja": "総合基本設計図", "ko": "대지 기본 설계도",
        "fr": "Plans d'avant-projet du site", "de": "Gelände-Vorprojektpläne",
        "es": "Planos preliminares del conjunto", "pt": "Desenhos preliminares do conjunto",
        "ru": "Эскизные чертежи генплана", "ar": "المخططات الأولية للموقع العام",
        "it": "Elaborati preliminari del sito", "vi": "Bản vẽ thiết kế cơ sở tổng mặt bằng",
        "th": "แบบผังรวมเบื้องต้น",
    },
    "建筑资料1（第一版作业图）": {
        "en": "Architectural Data 1 (1st Version)", "ja": "建築資料1（第一版作業図）", "ko": "건축 자료 1 (1차 작업도)",
        "fr": "Données architecturales 1 (1re version)", "de": "Architekturdaten 1 (1. Fassung)",
        "es": "Datos de arquitectura 1 (1.ª versión)", "pt": "Dados de arquitetura 1 (1.ª versão)",
        "ru": "Архитектурные данные 1 (первая версия)", "ar": "البيانات المعمارية 1 (النسخة الأولى)",
        "it": "Dati architettonici 1 (1ª versione)", "vi": "Dữ liệu kiến trúc 1 (lần 1)",
        "th": "ข้อมูลสถาปัตยกรรม 1 (ครั้งที่ 1)",
    },
    "建筑专业资料2": {
        "en": "Architectural Data 2", "ja": "建築資料2", "ko": "건축 자료 2",
        "fr": "Données architecturales 2", "de": "Architekturdaten 2",
        "es": "Datos de arquitectura 2", "pt": "Dados de arquitetura 2",
        "ru": "Архитектурные данные 2", "ar": "البيانات المعمارية 2",
        "it": "Dati architettonici 2", "vi": "Dữ liệu kiến trúc 2",
        "th": "ข้อมูลสถาปัตยกรรม 2",
    },
    "总图第一版作业图": {
        "en": "Site Working Drawings (1st Version)", "ja": "総合計画作業図（第一版）", "ko": "대지 작업도 (1차)",
        "fr": "Plans de travail du site (1re version)", "de": "Gelände-Arbeitspläne (1. Fassung)",
        "es": "Planos de trabajo del conjunto (1.ª versión)", "pt": "Desenhos de trabalho do conjunto (1.ª versão)",
        "ru": "Рабочие чертежи генплана (первая версия)", "ar": "مخططات العمل للموقع (النسخة الأولى)",
        "it": "Disegni di lavoro del sito (1ª versione)", "vi": "Bản vẽ làm việc tổng mặt bằng (lần 1)",
        "th": "แบบงานผังรวม (ครั้งที่ 1)",
    },
    "总图第二版作业图": {
        "en": "Site Working Drawings (2nd Version)", "ja": "総合計画作業図（第二版）", "ko": "대지 작업도 (2차)",
        "fr": "Plans de travail du site (2e version)", "de": "Gelände-Arbeitspläne (2. Fassung)",
        "es": "Planos de trabajo del conjunto (2.ª versión)", "pt": "Desenhos de trabalho do conjunto (2.ª versão)",
        "ru": "Рабочие чертежи генплана (вторая версия)", "ar": "مخططات العمل للموقع (النسخة الثانية)",
        "it": "Disegni di lavoro del sito (2ª versione)", "vi": "Bản vẽ làm việc tổng mặt bằng (lần 2)",
        "th": "แบบงานผังรวม (ครั้งที่ 2)",
    },
    "建筑资料2": {
        "en": "Architectural Data 2", "ja": "建築資料2", "ko": "건축 자료 2",
        "fr": "Données architecturales 2", "de": "Architekturdaten 2",
        "es": "Datos de arquitectura 2", "pt": "Dados de arquitetura 2",
        "ru": "Архитектурные данные 2", "ar": "البيانات المعمارية 2",
        "it": "Dati architettonici 2", "vi": "Dữ liệu kiến trúc 2",
        "th": "ข้อมูลสถาปัตยกรรม 2",
    },
    "水、暖专业提建筑、结构资料": {
        "en": "MEP Data to Architecture & Structure", "ja": "水・暖の建築・構造資料", "ko": "급배수·난방의 건축·구조 자료",
        "fr": "Données des réseaux vers architecture et structure", "de": "TGA-Daten an Architektur und Tragwerk",
        "es": "Datos de instalaciones a arquitectura y estructura", "pt": "Dados das instalações para arquitetura e estrutura",
        "ru": "Данные инженерных систем архитектуре и конструкторам", "ar": "بيانات الأنظمة إلى المعماري والإنشائي",
        "it": "Dati impianti ad architettura e struttura", "vi": "Dữ liệu hệ thống cho kiến trúc và kết cấu",
        "th": "ข้อมูลระบบวิศวกรรมให้สถาปัตยกรรมและโครงสร้าง",
    },
    "管道综合与对图协调会": {
        "en": "Pipe Coordination Meeting", "ja": "配管統合調整会議", "ko": "배관 종합 조정 회의",
        "fr": "Réunion de synthèse des réseaux", "de": "Rohrleitungskoordinationssitzung",
        "es": "Reunión de integración de redes", "pt": "Reunião de integração de redes",
        "ru": "Совещание по увязке трубопроводов", "ar": "اجتماع تنسيق الأنابيب",
        "it": "Riunione integrazione impianti", "vi": "Họp tổng hợp đường ống",
        "th": "ประชุมรวมระบบท่อ",
    },
    "水专业提清资料": {
        "en": "Water Supply Final Data", "ja": "給排水最終資料", "ko": "급배수 최종 자료",
        "fr": "Données finales des réseaux d'eau", "de": "Abschlussdaten Wasser",
        "es": "Datos finales de agua", "pt": "Dados finais de água",
        "ru": "Окончательные данные по водоснабжению", "ar": "البيانات النهائية لشبكات المياه",
        "it": "Dati finali idraulici", "vi": "Trình dữ liệu cuối cấp thoát nước",
        "th": "ส่งข้อมูลสุดท้ายระบบน้ำ",
    },
    "暖专业提清资料": {
        "en": "HVAC Final Data", "ja": "空調・暖房最終資料", "ko": "난방·공조 최종 자료",
        "fr": "Données finales HVAC", "de": "Abschlussdaten HLK",
        "es": "Datos finales de climatización", "pt": "Dados finais de HVAC",
        "ru": "Окончательные данные по отоплению и вентиляции", "ar": "البيانات النهائية للتدفئة والتهوية",
        "it": "Dati finali HVAC", "vi": "Trình dữ liệu cuối HVAC",
        "th": "ส่งข้อมูลสุดท้ายระบบปรับอากาศ",
    },
    "电专业提清资料": {
        "en": "Electrical Final Data", "ja": "電気最終資料", "ko": "전기 최종 자료",
        "fr": "Données finales électriques", "de": "Abschlussdaten Elektro",
        "es": "Datos finales eléctricos", "pt": "Dados finais elétricos",
        "ru": "Окончательные данные по электроснабжению", "ar": "البيانات النهائية الكهربائية",
        "it": "Dati finali elettrici", "vi": "Trình dữ liệu cuối điện",
        "th": "ส่งข้อมูลสุดท้ายระบบไฟฟ้า",
    },
    "智能化专业提清资料": {
        "en": "Intelligent Systems Final Data", "ja": "スマート設備最終資料", "ko": "지능형 설비 최종 자료",
        "fr": "Données finales des systèmes intelligents", "de": "Abschlussdaten Gebäudeautomation",
        "es": "Datos finales de sistemas inteligentes", "pt": "Dados finais de sistemas inteligentes",
        "ru": "Окончательные данные по интеллектуальным системам", "ar": "البيانات النهائية للأنظمة الذكية",
        "it": "Dati finali sistemi intelligenti", "vi": "Trình dữ liệu cuối hệ thống thông minh",
        "th": "ส่งข้อมูลสุดท้ายระบบอัจฉริยะ",
    },
    "建筑作业底图": {
        "en": "Architectural Base Drawings", "ja": "建築ベース図", "ko": "건축 베이스 도면",
        "fr": "Plan de base architectural", "de": "Architektonische Basisdaten",
        "es": "Plano base de arquitectura", "pt": "Desenho base de arquitetura",
        "ru": "Архитектурная основа для чертежей", "ar": "المخطط الأساسي المعماري",
        "it": "Base architettonica", "vi": "Bản vẽ nền kiến trúc",
        "th": "แบบพื้นฐานสถาปัตยกรรม",
    },
    "开始出图": {
        "en": "Start of Drawing Issue", "ja": "出図開始", "ko": "출도 시작",
        "fr": "Début de l'émission des plans", "de": "Beginn der Planausgabe",
        "es": "Inicio de emisión de planos", "pt": "Início da emissão de desenhos",
        "ru": "Начало выдачи чертежей", "ar": "بدء إصدار المخططات",
        "it": "Inizio emissione disegni", "vi": "Bắt đầu phát bản vẽ",
        "th": "เริ่มส่งแบบ",
    },
    "交付": {
        "en": "Delivery", "ja": "納品", "ko": "인도",
        "fr": "Livraison", "de": "Übergabe", "es": "Entrega", "pt": "Entrega",
        "ru": "Сдача", "ar": "التسليم", "it": "Consegna", "vi": "Bàn giao",
        "th": "ส่งมอบ",
    },
    # ---- 专业名（专业作业清单表名/行内）----
    "建筑": {
        "en": "Architecture", "ja": "建築", "ko": "건축",
        "fr": "Architecture", "de": "Architektur", "es": "Arquitectura",
        "pt": "Arquitetura", "ru": "Архитектура", "ar": "المعمارية",
        "it": "Architettura", "vi": "Kiến trúc", "th": "สถาปัตยกรรม",
    },
    "总平面": {
        "en": "Site Planning", "ja": "総合計画", "ko": "대지 계획",
        "fr": "Plan d'ensemble", "de": "Geländeplanung", "es": "Planeamiento del conjunto",
        "pt": "Planejamento do conjunto", "ru": "Генеральный план", "ar": "التخطيط العام",
        "it": "Pianificazione del sito", "vi": "Quy hoạch tổng mặt bằng", "th": "การวางผัง",
    },
    "结构": {
        "en": "Structure", "ja": "構造", "ko": "구조",
        "fr": "Structure", "de": "Tragwerk", "es": "Estructura", "pt": "Estrutura",
        "ru": "Конструкции", "ar": "الإنشائي", "it": "Struttura", "vi": "Kết cấu",
        "th": "โครงสร้าง",
    },
    "给排水": {
        "en": "Water Supply & Drainage", "ja": "給排水", "ko": "급배수",
        "fr": "Alimentation en eau et assainissement", "de": "Wasser/Abwasser",
        "es": "Abastecimiento y drenaje", "pt": "Abastecimento e drenagem",
        "ru": "Водоснабжение и канализация", "ar": "إمدادات المياه والصرف",
        "it": "Idraulica", "vi": "Cấp thoát nước", "th": "ระบบน้ำและระบายน้ำ",
    },
    "暖通": {
        "en": "HVAC", "ja": "空調・暖房", "ko": "난방·공조",
        "fr": "CVC", "de": "HLK", "es": "Climatización", "pt": "AVAC",
        "ru": "ОВиК", "ar": "التدفئة والتهوية والتكييف", "it": "Riscaldamento e ventilazione",
        "vi": "HVAC", "th": "ระบบปรับอากาศ",
    },
    "电气": {
        "en": "Electrical", "ja": "電気", "ko": "전기",
        "fr": "Électricité", "de": "Elektro", "es": "Eléctrico", "pt": "Elétrica",
        "ru": "Электроснабжение", "ar": "الكهرباء", "it": "Elettrico", "vi": "Điện",
        "th": "ระบบไฟฟ้า",
    },
    "智能化": {
        "en": "Intelligent Systems", "ja": "スマート設備", "ko": "지능형 설비",
        "fr": "Systèmes intelligents", "de": "Gebäudeautomation", "es": "Sistemas inteligentes",
        "pt": "Sistemas inteligentes", "ru": "Интеллектуальные системы", "ar": "الأنظمة الذكية",
        "it": "Sistemi intelligenti", "vi": "Hệ thống thông minh", "th": "ระบบอัจฉริยะ",
    },
    "经济": {
        "en": "Economics", "ja": "経済", "ko": "경제",
        "fr": "Économie", "de": "Wirtschaftlichkeit", "es": "Economía", "pt": "Economia",
        "ru": "Экономика", "ar": "الاقتصاد", "it": "Economia", "vi": "Kinh tế",
        "th": "เศรษฐศาสตร์",
    },
}

# =====================================================================
# v5.0 扩展：节点"具体内容/事项"条目级译文映射
# ---------------------------------------------------------------------
# 基于《建筑工程设计项目专业配合规定》(2015版) 源文档的节点原文建立
# 中 -> 各语言映射；仅收录高频/通用条目（含两阶段共用的协调会组织句、
# 图纸归档句、会审会签句、结构/机电反馈意见句等），每条覆盖全部 12 种
# 语言（en/ja/ko/fr/de/es/pt/ru/ar/it/vi/th）。
# 未收录的条目保持中文原文，严禁编造或机翻式硬译（宁缺勿滥）。
# 呈现方式与 LANG_DICT 一致：中文 + [CODE] 译文依次追加。
# =====================================================================
CONTENT_TRANS = {
    "项目经理组织、设计主持人主持、各专业负责人参加：介绍项目基本情况，发放《设计任务书》和《设计策划表》；发放经甲方确认的方案图；研究和讨论甲方及政府部门要求，制定统一设计要求和措施；制定本工程《专业配合进度表》并签字生效": {
        "en": "Organized by the project manager, chaired by the design director, attended by the heads of all disciplines: introduce the basic project information and issue the Design Task Book and the Design Planning Sheet; issue the scheme drawings confirmed by the client; study and discuss the requirements of the client and government authorities, and formulate unified design requirements and measures; prepare the Professional Coordination Schedule for this project and sign it into effect",
        "ja": "プロジェクトマネージャーが組織し、設計主催者が司会し、各専門担当責任者が参加する：プロジェクトの基本情報を紹介し、『設計任務書』と『設計計画表』を配布する；発注者が確認した方案図を配布する；発注者および政府部門の要求を研究・検討し、統一した設計要求と措置を策定する；本工事の『専門配合進度表』を作成し、署名して発効させる",
        "ko": "프로젝트 매니저가 조직하고 설계 책임자가 주재하며 각 전문 담당자가 참석합니다: 프로젝트 기본 정보를 소개하고 《설계임무서》와 《설계기획표》를 배포합니다; 발주자가 확인한 안(案) 도면을 배포합니다; 발주자와 정부 부처의 요구를 연구·토론하고 통일된 설계 요구와 조치를 수립합니다; 본 공사의 《전문배합진도표》를 작성하여 서명 후 효력이 발생합니다",
        "fr": "Organisé par le chef de projet, présidé par le directeur de conception, avec la participation des responsables de toutes les disciplines : présentation des informations générales du projet et distribution du Cahier des missions de conception et de la Fiche de planification de conception ; distribution des plans de projet confirmés par le maître d'ouvrage ; étude et discussion des exigences du maître d'ouvrage et des autorités publiques, formulation d'exigences et de mesures de conception unifiées ; établissement du Calendrier de coordination professionnelle du présent projet, signé et entré en vigueur",
        "de": "Organisiert durch den Projektleiter, geleitet durch den Entwurfsleiter, mit Teilnahme der Verantwortlichen aller Fachdisziplinen: Vorstellung der grundlegenden Projektdaten und Ausgabe des Entwurfsauftragsbuchs und des Entwurfsplanungsblatts; Ausgabe der vom Bauherrn bestätigten Planungszeichnungen; Untersuchung und Diskussion der Anforderungen des Bauherrn und der Behörden sowie Festlegung einheitlicher Entwurfsanforderungen und -maßnahmen; Erstellung des Fachkoordinationszeitplans für dieses Bauvorhaben, Unterzeichnung und Inkraftsetzung",
        "es": "Organizado por el jefe de proyecto, presidido por el director de diseño, con participación de los responsables de todas las especialidades: presentación de la información básica del proyecto y entrega del Encargo de diseño y de la Hoja de planificación del diseño; entrega de los planos de esquema confirmados por el cliente; estudio y discusión de los requisitos del cliente y de las autoridades públicas, formulación de requisitos y medidas de diseño unificados; elaboración del Cronograma de coordinación profesional de esta obra, firmado y en vigor",
        "pt": "Organizado pelo gerente do projeto, presidido pelo diretor de projeto, com a participação dos responsáveis de todas as especialidades: apresentação das informações básicas do projeto e entrega do Termo de Tarefa de Projeto e da Ficha de Planejamento de Projeto; entrega dos desenhos de estudo confirmados pelo cliente; estudo e discussão dos requisitos do cliente e dos órgãos públicos, formulação de requisitos e medidas unificados de projeto; elaboração do Cronograma de Coordenação Profissional desta obra, assinado e em vigor",
        "ru": "Организует руководитель проекта, ведёт главный архитектор проекта, участвуют ответственные по всем специальностям: представление основных сведений о проекте и выдача Задания на проектирование и Плана проектирования; выдача утверждённых заказчиком эскизных чертежей; изучение и обсуждение требований заказчика и государственных органов, разработка единых требований и мероприятий по проектированию; составление Графика профессиональной координации по данному объекту с подписанием и вступлением в силу",
        "ar": "ينظمه مدير المشروع، ويرأسه مدير التصميم، ويشارك فيه رؤساء جميع التخصصات: تقديم المعلومات الأساسية عن المشروع وتوزيع كتاب مهام التصميم ونموذج تخطيط التصميم؛ وتوزيع المخططات التصميمية المعتمدة من العميل؛ ودراسة ومناقشة متطلبات العميل والجهات الحكومية ووضع متطلبات وإجراءات تصميم موحدة؛ وإعداد جدول التنسيق المهني لهذا المشروع وتوقيعه ودخوله حيز التنفيذ",
        "it": "Organizzato dal project manager, presieduto dal direttore della progettazione, con la partecipazione dei responsabili di tutte le discipline: presentazione delle informazioni generali del progetto e distribuzione del Libretto dei compiti di progettazione e del Foglio di pianificazione della progettazione; distribuzione degli elaborati schematici confermati dal committente; studio e discussione dei requisiti del committente e delle autorità pubbliche, definizione di requisiti e misure progettuali unificati; predisposizione del Cronoprogramma di coordinamento professionale di questo progetto, firmato ed entrato in vigore",
        "vi": "Do quản lý dự án tổ chức, giám đốc thiết kế chủ trì, các trưởng bộ môn tham dự: giới thiệu thông tin cơ bản của dự án, phát hành Sổ nhiệm vụ thiết kế và Phiếu kế hoạch thiết kế; phát hành bản vẽ phương án đã được chủ đầu tư xác nhận; nghiên cứu thảo luận các yêu cầu của chủ đầu tư và cơ quan nhà nước, đề ra các yêu cầu và biện pháp thiết kế thống nhất; lập Bảng tiến độ phối hợp chuyên môn của công trình này và ký để có hiệu lực",
        "th": "จัดโดยผู้จัดการโครงการ โดยผู้อำนวยการฝ่ายออกแบบเป็นประธาน และหัวหน้าทุกสาขาวิชาเข้าร่วม: แนะนำข้อมูลพื้นฐานของโครงการ แจกหนังสือภารกิจออกแบบและแบบฟอร์มวางแผนการออกแบบ; แจกแบบสถาปัตยกรรมที่ได้รับยืนยันจากเจ้าของโครงการ; ศึกษาและอภิปรายข้อกำหนดของเจ้าของโครงการและหน่วยงานภาครัฐ กำหนดข้อกำหนดและมาตรการออกแบบที่สอดคล้องกัน; จัดทำตารางการประสานงานวิชาชีพของโครงการนี้และลงนามให้มีผลบังคับใช้",
    },
    "建筑专业提出，用于各专业初步设计准备图。内容：方案设计说明；全部方案图纸（平面图、立面图、剖面图，深度满足建设部对方案设计深度的要求）；政府、甲方或业主对方案的修改和（或）审批意见；其他需要提出的资料。接受专业：总图、结构、水、暖、电、智能化、经济": {
        "en": "Proposed by the architectural discipline as preparation drawings for the preliminary design of all disciplines. Contents: description of the scheme design; all scheme drawings (plans, elevations and sections, with depth satisfying the requirements of the Ministry of Construction on scheme design depth); comments of the government, the client or the owner on modification and/or approval of the scheme; other materials to be provided. Receiving disciplines: site planning, structural, plumbing, HVAC, electrical, intelligent systems, economics",
        "ja": "建築専門が提出し、各専門の基本設計準備図に供する。内容：方案設計説明；全方案図面（平面図、立面図、断面図。深さは建設部の方案設計深度の要求を満たす）；政府・発注者・オーナーによる方案の修正および（または）承認意見；その他提出すべき資料。受取専門：総図、構造、水、暖、電、智能化、経済",
        "ko": "건축 전문이 제출하며 각 전문의 기본설계 준비도면으로 사용됩니다. 내용: 안(案) 설계 설명; 전체 안 도면(평면도, 입면도, 단면도, 깊이는 건설부의 안 설계 깊이 요구를 충족); 정부·발주자·소유자의 안 수정 및(또는) 승인 의견; 기타 제출해야 할 자료. 수신 전문: 총도, 구조, 급배수, 난방, 전기, 지능화, 경제",
        "fr": "Présenté par la discipline architecture, servant de plans préparatoires au projet d'esquisse pour toutes les disciplines. Contenu : notice de conception du schéma ; l'ensemble des plans du schéma (plans, élévations, coupes, avec une profondeur conforme aux exigences du Ministère de la Construction) ; avis de modification et/ou d'approbation du schéma émis par les pouvoirs publics, le maître d'ouvrage ou le propriétaire ; autres documents à fournir. Disciplines destinataires : aménagement du site, structure, plomberie, CVC, électricité, systèmes intelligents, économie",
        "de": "Vorgelegt von der Architekturdisziplin als Vorbereitungszeichnungen für die Vorplanung aller Fachdisziplinen. Inhalt: Beschreibung der Entwurfskonzeption; sämtliche Konzeptzeichnungen (Grundrisse, Ansichten, Schnitte, deren Tiefe den Anforderungen des Bauministeriums an die Konzepttiefe genügt); Änderungs- und/oder Genehmigungsanmerkungen der Behörden, des Bauherrn oder des Eigentümers zum Konzept; sonstige vorzulegende Unterlagen. Empfangende Disziplinen: Standortplanung, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de arquitectura, como planos preparatorios para el diseño preliminar de todas las especialidades. Contenido: memoria del diseño esquemático; todos los planos del esquema (plantas, alzados y secciones, con profundidad conforme a los requisitos del Ministerio de Construcción sobre profundidad del diseño esquemático); observaciones del gobierno, del cliente o del propietario sobre la modificación y/o aprobación del esquema; demás documentación a aportar. Especialidades receptoras: urbanización, estructuras, fontanería, climatización, electricidad, sistemas inteligentes, economía",
        "pt": "Apresentado pela especialidade de arquitetura, como desenhos preparatórios para o projeto preliminar de todas as especialidades. Conteúdo: memorial do projeto de estudo; todos os desenhos do estudo (plantas, fachadas e cortes, com profundidade atendendo aos requisitos do Ministério da Construção sobre a profundidade do projeto de estudo); observações do governo, do cliente ou do proprietário sobre a modificação e/ou aprovação do estudo; demais documentos a fornecer. Especialidades receptoras: urbanismo, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes, economia",
        "ru": "Представляется архитектурной специальностью в качестве подготовительных чертежей для разработки стадии «проект» всеми специальностями. Содержание: пояснительная записка к архитектурному решению; все чертежи архитектурного решения (планы, фасады, разрезы, с глубиной, отвечающей требованиям Министерства строительства к глубине проработки эскизного проекта); замечания государственных органов, заказчика или собственника по изменению и/или утверждению решения; прочие материалы, подлежащие представлению. Принимающие специальности: генплан, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы, экономика",
        "ar": "يقدمه تخصص العمارة كمخططات تحضيرية للتصميم الأولي لجميع التخصصات. المحتوى: وصف التصميم المبدئي؛ وجميع مخططات التصميم المبدئي (مخططات أفقية وواجهات وقطاعات بعمق يفي بمتطلبات وزارة الإنشاءات بشأن عمق التصميم المبدئي)؛ وملاحظات الحكومة أو العميل أو المالك بشأن تعديل و/أو اعتماد التصميم المبدئي؛ وغيرها من المواد التي يلزم تقديمها. التخصصات المستلمة: تخطيط الموقع، الإنشاءات، السباكة، التدفئة والتهوية، الكهرباء، الأنظمة الذكية، الاقتصاد",
        "it": "Presentato dalla disciplina architettonica come elaborati preparatori per la progettazione preliminare di tutte le discipline. Contenuto: relazione del progetto preliminare; tutti gli elaborati del progetto preliminare (piante, prospetti e sezioni, con profondità conforme ai requisiti del Ministero delle Costruzioni sulla profondità del progetto preliminare); osservazioni del governo, del committente o del proprietario sulla modifica e/o approvazione del progetto preliminare; altri documenti da fornire. Discipline destinatarie: pianificazione del sito, strutture, impianti idrici, HVAC, elettrici, sistemi intelligenti, economia",
        "vi": "Do chuyên ngành kiến trúc đưa ra, dùng làm bản vẽ chuẩn bị cho thiết kế cơ sở của các chuyên ngành. Nội dung: thuyết minh thiết kế phương án; toàn bộ bản vẽ phương án (mặt bằng, mặt đứng, mặt cắt, độ sâu đáp ứng yêu cầu của Bộ Xây dựng về độ sâu thiết kế phương án); ý kiến của chính phủ, chủ đầu tư hoặc chủ sở hữu về việc sửa đổi và/hoặc phê duyệt phương án; các tài liệu khác cần nêu. Các chuyên ngành nhận: tổng mặt bằng, kết cấu, cấp thoát nước, điều hòa không khí, điện, hệ thống thông minh, kinh tế",
        "th": "เสนอโดยสาขาสถาปัตยกรรม ใช้เป็นแบบเตรียมการออกแบบขั้นต้นของทุกสาขา เนื้อหา: คำอธิบายการออกแบบแบบสถาปัตยกรรม; แบบสถาปัตยกรรมทั้งหมด (แบบพื้น แบบด้าน แบบตัด ให้ความลึกรายละเอียดเป็นไปตามข้อกำหนดของกระทรวงการก่อสร้างว่าด้วยความลึกของการออกแบบแบบสถาปัตยกรรม); ความเห็นของภาครัฐ เจ้าของโครงการ หรือเจ้าของทรัพย์สิน ในการแก้ไขและ/หรืออนุมัติแบบสถาปัตยกรรม; เอกสารอื่นที่ต้องเสนอ สาขาที่รับ: การวางผังบริเวณ โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ เศรษฐศาสตร์",
    },
    "结构专业提给建筑专业：考虑结构方案、研究建筑方案是否满足结构可行性；提出地质勘察要求；提出结构构造措施（伸缩缝、沉降缝、防震缝位置要求及特殊部位构造做法）；提出人工地基处理方法、处理深度；提出结构选型（砖混、框架、框剪、筒体、转换结构等）；提出基础形式（独立、条形、十字交叉梁、筏基、桩基、箱型等）及基础截面估算尺寸；对建筑方案修改的其他建议": {
        "en": "The structure discipline submits to the architecture discipline: consider the structural scheme and study whether the architectural scheme meets structural feasibility; propose requirements for geological investigation; propose structural detailing measures (requirements for positions of expansion joints, settlement joints, seismic joints and construction practices at special parts); propose methods and depth of artificial ground treatment; propose structural system selection (masonry-concrete, frame, frame-shear wall, tube, transfer structure, etc.); propose foundation types (independent, strip, cross beam raft, raft, pile, box, etc.) and estimated foundation section dimensions; other suggestions for modification of the architectural scheme",
        "ja": "構造専門が建築専門に提出する：構造案を検討し、建築案が構造の実現可能性を満たすかを研究する；地質調査の要求を提出する；構造措置（伸縮継目、沈下継目、耐震継目の位置要求と特殊部位の構造作法）を提出する；人工地盤処理方法、処理深さを提出する；構造選型（組積造・コンクリート、フレーム、フレーム・耐震壁、チューブ、変換構造等）を提出する；基礎形式（独立、ベタ基礎、十字交差梁、筏式、杭基礎、箱型等）及び基礎断面の概算寸法を提出する；建築案修正に関するその他の提案",
        "ko": "구조 전문이 건축 전문에 제출한다：구조안을 고려하고 건축안이 구조 실현 가능성을 충족하는지 연구한다；지질 조사 요구를 제출한다；구조 조치(신축이음, 침하이음, 내진이음 위치 요구 및 특수 부위 시공법)를 제출한다；인공 지반 처리 방법, 처리 깊이를 제출한다；구조 선형(조적-콘크리트, 프레임, 프레임-전단벽, 튜브, 전환 구조 등)을 제출한다；기초 형식(독립, 줄기초, 십자교차보, 전면기초(筏基), 파일기초, 상자형 등) 및 기초 단면 추정 치수를 제출한다；건축안 수정에 관한 기타 제안",
        "fr": "La spécialité structure soumet à la spécialité d'architecture : considérer le schéma structurel et étudier si le schéma architectural répond à la faisabilité structurelle ; proposer les exigences d'étude géologique ; proposer les mesures constructives (exigences de position des joints de dilatation, des joints de tassement, des joints sismiques et pratiques constructives des parties spéciales) ; proposer les méthodes et la profondeur de traitement du sol artificiel ; proposer la sélection du système structurel (maçonnerie-béton, portique, portique-voile, tube, structure de transition, etc.) ; proposer les types de fondations (isolées, filantes, croisillons de poutres, radier, pieux, caisson, etc.) et les dimensions estimées des sections de fondations ; autres suggestions de modification du schéma architectural",
        "de": "Die Fachrichtung Tragwerksplanung legt der Fachrichtung Architektur vor: das Tragwerksschema prüfen und untersuchen, ob das Architekturschema die Machbarkeit des Tragwerks erfüllt; Anforderungen an die geologische Untersuchung stellen; konstruktive Maßnahmen vorschlagen (Anforderungen an die Positionen von Dehnungsfugen, Setzungsfugen, Erdbebenfugen und konstruktive Ausführungen an besonderen Stellen); Methoden und Tiefe der künstlichen Baugrundbehandlung vorschlagen; die Auswahl des Tragsystems vorschlagen (Mauerwerk-Beton, Rahmen, Rahmen-Wandscheiben, Röhre, Übergangskonstruktion usw.); Gründungsarten vorschlagen (Einzelfundamente, Streifenfundamente, Kreuzbalken, Plattengründung, Pfähle, Kasten usw.) sowie geschätzte Abmessungen der Gründungsquerschnitte; sonstige Vorschläge zur Änderung des Architekturschemas",
        "es": "La especialidad de estructura presenta a la especialidad de arquitectura: considerar el esquema estructural y estudiar si el esquema arquitectónico cumple la viabilidad estructural; proponer los requisitos de estudio geológico; proponer las medidas constructivas (requisitos de posición de las juntas de dilatación, juntas de asentamiento, juntas sísmicas y prácticas constructivas en partes especiales); proponer los métodos y la profundidad del tratamiento del terreno artificial; proponer la selección del sistema estructural (mampostería-hormigón, pórtico, pórtico-muro, tubo, estructura de transición, etc.); proponer los tipos de cimentación (aislada, corrida, emparrillado, losa, pilotes, cajón, etc.) y las dimensiones estimadas de las secciones de cimentación; otras sugerencias de modificación del esquema arquitectónico",
        "pt": "A especialidade estrutural apresenta à especialidade de arquitetura: considerar o esquema estrutural e estudar se o esquema arquitetônico atende à viabilidade estrutural; propor os requisitos de investigação geológica; propor as medidas construtivas (requisitos de posição das juntas de dilatação, juntas de assentamento, juntas sísmicas e práticas construtivas em partes especiais); propor os métodos e a profundidade do tratamento do terreno artificial; propor a seleção do sistema estrutural (alvenaria-concreto, pórtico, pórtico-parede, tubo, estrutura de transição, etc.); propor os tipos de fundação (isolada, corrida, grelha, radier, estacas, caixão, etc.) e as dimensões estimadas das seções de fundação; outras sugestões de modificação do esquema arquitetônico",
        "ru": "Строительная специальность представляет архитектурной специальности: рассмотреть конструктивную схему и изучить, отвечает ли архитектурная схема требованиям осуществимости конструкции; выдвинуть требования к инженерно-геологическим изысканиям; предложить конструктивные мероприятия (требования к расположению деформационных, осадочных, сейсмических швов и конструктивные решения особых мест); предложить методы и глубину искусственного улучшения основания; предложить выбор конструктивной системы (кирпично-монолитная, каркасная, каркасно-стеновая, трубная, переходная конструкция и т.п.); предложить типы фундаментов (отдельные, ленточные, перекрестные, плитные, свайные, коробчатые и т.п.) и ориентировочные размеры сечений фундаментов; прочие предложения по изменению архитектурной схемы",
        "ar": "يقدم تخصص الإنشاءات إلى تخصص العمارة: النظر في المخطط الإنشائي ودراسة ما إذا كان المخطط المعماري يلبي الجدوى الإنشائية؛ اقتراح متطلبات التحقيق الجيولوجي؛ اقتراح التدابير الإنشائية (متطلبات مواقع فواصل التمدد وفواصل الهبوط وفواصل الزلازل وأساليب الإنشاء في الأجزاء الخاصة)؛ اقتراح طرق وعمق معالجة التربة الاصطناعية؛ اقتراح اختيار النظام الإنشائي (البناء بالطوب والخرسانة، الإطار، الإطار والجدران القصية، الأنبوبي، البنية الانتقالية وغيرها)؛ اقتراح أنواع الأساسات (المنفردة، الشريطية، الشبكية، اللوحية، الخوازيق، الصندوقية وغيرها) والأبعاد التقديرية لمقاطع الأساسات؛ اقتراحات أخرى لتعديل المخطط المعماري",
        "it": "La disciplina strutturale presenta alla disciplina architettonica: considerare lo schema strutturale e studiare se lo schema architettonico soddisfa la fattibilità strutturale; proporre i requisiti delle indagini geologiche; proporre i provvedimenti costruttivi (requisiti di posizione dei giunti di dilatazione, giunti di assestamento, giunti sismici e modalità costruttive nelle parti speciali); proporre i metodi e la profondità del trattamento artificiale del terreno; proporre la selezione del sistema strutturale (muratura-cemento, telaio, telaio-pareti, tubolare, struttura di transizione, ecc.); proporre i tipi di fondazione (isolata, continua, a graticcio, a platea, su pali, a cassone, ecc.) e le dimensioni stimate delle sezioni di fondazione; altri suggerimenti per la modifica dello schema architettonico",
        "vi": "Chuyên ngành kết cấu trình cho chuyên ngành kiến trúc: xem xét phương án kết cấu và nghiên cứu phương án kiến trúc có đáp ứng tính khả thi về kết cấu hay không; đề xuất yêu cầu khảo sát địa chất; đề xuất các biện pháp cấu tạo (yêu cầu vị trí khe co giãn, khe lún, khe chống động đất và cách thi công các bộ phận đặc biệt); đề xuất phương pháp và chiều sâu xử lý nền nhân tạo; đề xuất lựa chọn hệ kết cấu (gạch-bê tông, khung, khung-vách, ống, kết cấu chuyển đổi...); đề xuất hình thức móng (móng đơn, móng băng, móng giao thoa, móng bè, móng cọc, móng hộp...) và kích thước ước tính mặt cắt móng; các đề xuất khác về sửa đổi phương án kiến trúc",
        "th": "สาขาโครงสร้างยื่นต่อสาขาสถาปัตยกรรม: พิจารณาแผนโครงสร้างและศึกษาแผนสถาปัตยกรรมว่าตอบสนองความเป็นไปได้ด้านโครงสร้างหรือไม่ เสนอข้อกำหนดการสำรวจธรณีวิทยา เสนอมาตรการก่อสร้าง (ข้อกำหนดตำแหน่งรอยต่อขยายตัว รอยต่อทรุดตัว รอยต่อต้านแผ่นดินไหว และวิธีก่อสร้างส่วนพิเศษ) เสนอวิธีการและความลึกของการปรับปรุงฐานรากเทียม เสนอการเลือกระบบโครงสร้าง (ก่ออิฐ-คอนกรีต โครง โครง-ผนังเฉือน กล่อง โครงสร้างถ่ายน้ำหนัก ฯลฯ) เสนอรูปแบบฐานราก (ฐานเดี่ยว ฐานแถบ ฐานคร่อม ฐานแผ่ ฐานเสาเข็ม ฐานกล่อง ฯลฯ) และขนาดโดยประมาณของหน้าตัดฐานราก ข้อเสนออื่น ๆ สำหรับการปรับแผนสถาปัตยกรรม",
    },
    "机电专业提给建筑专业：考虑机电系统设计方案、研究建筑方案是否满足机电系统可行性；提出机电系统能源站房大致面积、建议位置或区域；提出各分散机房大致面积、位置的确定原则；对主要大型管道井提出设置原则和面积要求；大型设备吊装孔位置和尺寸要求及运输通道要求；对建筑方案修改的其他建议": {
        "en": "The MEP discipline submits to the architecture discipline: consider the design scheme of MEP systems and study whether the architectural scheme meets the feasibility of MEP systems; propose the approximate area and suggested location or area of the MEP energy station rooms; propose the principles for determining the approximate area and location of each decentralized equipment room; propose setting principles and area requirements for the main large pipe shafts; requirements for the position and dimensions of lifting openings for large equipment and for transport passages; other suggestions for modification of the architectural scheme",
        "ja": "机电専門が建築専門に提出する：机电システムの設計案を検討し、建築案が机电システムの実現可能性を満たすかを研究する；机电システム能源站房の概略面積、推奨位置又は区域を提出する；各分散機械室の概略面積、位置の決定原則を提出する；主要な大型パイプシャフトに設置原則と面積要求を提出する；大型設備の吊り上げ孔の位置・寸法要求と運搬通路要求；建築案修正に関するその他の提案",
        "ko": "기전 전문이 건축 전문에 제출한다：기전 시스템 설계안을 고려하고 건축안이 기전 시스템 실현 가능성을 충족하는지 연구한다；기전 시스템 에너지 스테이션 실의 대략 면적, 권장 위치 또는 구역을 제시한다；각 분산 기계실의 대략 면적·위치 결정 원칙을 제시한다；주요 대형 파이프 샤프트에 설치 원칙과 면적 요구를 제시한다；대형 설비 인양구 위치·치수 요구와 운반 통로 요구；건축안 수정에 관한 기타 제안",
        "fr": "La spécialité CVC-électricité-plomberie soumet à la spécialité d'architecture : examiner le schéma de conception des systèmes CVC-électricité-plomberie et étudier si le schéma architectural répond à la faisabilité des systèmes ; proposer la superficie approximative et l'emplacement ou la zone recommandés des salles de stations d'énergie des systèmes ; proposer les principes de détermination de la superficie approximative et de l'emplacement de chaque salle d'équipements décentralisée ; proposer les principes d'implantation et les exigences de superficie pour les principales gaines techniques de grande dimension ; exigences relatives à la position et aux dimensions des ouvertures de levage pour les grands équipements et aux voies de transport ; autres suggestions de modification du schéma architectural",
        "de": "Die Fachrichtung Technische Gebäudeausrüstung legt der Fachrichtung Architektur vor: das Entwurfsschema der TGA-Systeme prüfen und untersuchen, ob das Architekturschema die Machbarkeit der TGA-Systeme erfüllt; die ungefähre Fläche und den empfohlenen Standort oder Bereich der Energieversorgungsräume der TGA-Systeme vorschlagen; die Grundsätze zur Bestimmung der ungefähren Fläche und des Standorts der einzelnen dezentralen Technikräume vorschlagen; für die wichtigsten großen Installationsschächte Einrichtungsgrundsätze und Flächenanforderungen vorschlagen; Anforderungen an Position und Abmessungen der Montageöffnungen für große Anlagen sowie an Transportwege; sonstige Vorschläge zur Änderung des Architekturschemas",
        "es": "La especialidad de instalaciones electromecánicas presenta a la especialidad de arquitectura: considerar el esquema de diseño de los sistemas electromecánicos y estudiar si el esquema arquitectónico cumple la viabilidad de los sistemas electromecánicos; proponer la superficie aproximada y la ubicación o zona recomendadas de las salas de estaciones de energía de los sistemas electromecánicos; proponer los principios para determinar la superficie aproximada y la ubicación de cada sala de equipos descentralizada; proponer principios de implantación y requisitos de superficie para los principales patinillos de gran dimensión; requisitos de posición y dimensiones de las aberturas de izado para grandes equipos y de vías de transporte; otras sugerencias de modificación del esquema arquitectónico",
        "pt": "A especialidade de instalações eletromecânicas apresenta à especialidade de arquitetura: considerar o esquema de projeto dos sistemas eletromecânicos e estudar se o esquema arquitetônico atende à viabilidade dos sistemas eletromecânicos; propor a área aproximada e a localização ou zona recomendadas das salas das centrais de energia dos sistemas eletromecânicos; propor os princípios para determinar a área aproximada e a localização de cada sala de equipamentos descentralizada; propor princípios de implantação e requisitos de área para os principais shafts de grande porte; requisitos de posição e dimensões das aberturas de içamento para grandes equipamentos e de vias de transporte; outras sugestões de modificação do esquema arquitetônico",
        "ru": "Специальность инженерных систем здания представляет архитектурной специальности: рассмотреть проектную схему инженерных систем и изучить, отвечает ли архитектурная схема требованиям осуществимости инженерных систем; предложить ориентировочную площадь и рекомендуемое размещение или зону помещений энергетических станций инженерных систем; предложить принципы определения ориентировочной площади и размещения каждого децентрализованного технического помещения; для основных крупных инженерных шахт предложить принципы устройства и требования к площади; требования к положению и размерам монтажных проемов для крупного оборудования и к транспортным путям; прочие предложения по изменению архитектурной схемы",
        "ar": "يقدم تخصص الأنظمة الميكانيكية والكهربائية والسباكة إلى تخصص العمارة: النظر في مخطط تصميم الأنظمة ودراسة ما إذا كان المخطط المعماري يلبي جدوى الأنظمة؛ اقتراح المساحة التقريبية والموقع أو المنطقة الموصى بها لغرف محطات الطاقة للأنظمة؛ اقتراح مبادئ تحديد المساحة والموقع التقريبيين لكل غرفة معدات موزعة؛ اقتراح مبادئ الإنشاء ومتطلبات المساحة لأعمدة الخدمات الرئيسية الكبيرة؛ متطلبات موقع وأبعاد فتحات الرفع للمعدات الكبيرة وممرات النقل؛ اقتراحات أخرى لتعديل المخطط المعماري",
        "it": "La disciplina impianti meccanici-elettrici-idraulici presenta alla disciplina architettonica: considerare lo schema progettuale degli impianti e studiare se lo schema architettonico soddisfa la fattibilità degli impianti; proporre l'area approssimativa e la posizione o zona consigliate delle sale delle centrali energetiche degli impianti; proporre i principi per determinare l'area approssimativa e la posizione di ogni locale tecnico decentrato; proporre principi di installazione e requisiti di area per i principali grandi cavedi impiantistici; requisiti di posizione e dimensioni delle aperture di sollevamento per grandi apparecchiature e dei percorsi di trasporto; altri suggerimenti per la modifica dello schema architettonico",
        "vi": "Chuyên ngành cơ điện trình cho chuyên ngành kiến trúc: xem xét phương án thiết kế hệ thống cơ điện và nghiên cứu phương án kiến trúc có đáp ứng tính khả thi của hệ thống cơ điện hay không; đề xuất diện tích gần đúng và vị trí hoặc khu vực đề xuất của phòng trạm năng lượng hệ thống cơ điện; đề xuất nguyên tắc xác định diện tích gần đúng và vị trí của từng phòng máy phân tán; đề xuất nguyên tắc bố trí và yêu cầu diện tích cho các hộp kỹ thuật lớn chủ yếu; yêu cầu về vị trí, kích thước lỗ cẩu thiết bị lớn và lối vận chuyển; các đề xuất khác về sửa đổi phương án kiến trúc",
        "th": "สาขาเครื่องกลไฟฟ้าประปายื่นต่อสาขาสถาปัตยกรรม: พิจารณาแผนการออกแบบระบบเครื่องกลไฟฟ้าและศึกษาแผนสถาปัตยกรรมว่าตอบสนองความเป็นไปได้ของระบบเครื่องกลไฟฟ้าหรือไม่ เสนอพื้นที่โดยประมาณและตำแหน่งหรือเขตที่แนะนำของห้องสถานีพลังงานของระบบเครื่องกลไฟฟ้า เสนอหลักการกำหนดพื้นที่โดยประมาณและตำแหน่งของห้องเครื่องที่กระจายตัวแต่ละห้อง เสนอหลักการติดตั้งและข้อกำหนดด้านพื้นที่สำหรับปล่องบริการขนาดใหญ่หลัก ข้อกำหนดตำแหน่งและขนาดของช่องยกอุปกรณ์ขนาดใหญ่และเส้นทางขนส่ง ข้อเสนออื่น ๆ สำหรับการปรับแผนสถาปัตยกรรม",
    },
    "总图专业提出。内容：场地总平面（场地红线及红线内已有建筑、构筑物、道路、绿化等的名称及布置情况）；简要市政条件（可采用文字说明）；用地红线；其他需要提出的资料。接收专业：建筑、结构、水、暖、电、智能化、经济": {
        "en": "Proposed by the site planning discipline. Contents: general site plan (the site boundary and the names and layout of existing buildings, structures, roads, landscaping, etc. within the boundary); brief municipal conditions (may be given in text); land use boundary; other materials to be provided. Receiving disciplines: architectural, structural, plumbing, HVAC, electrical, intelligent systems, economics",
        "ja": "総図専門が提出する。内容：敷地全体計画（敷地境界線および境界線内の既存建築物・構築物・道路・緑化等の名称と配置状況）；簡潔な都市インフラ条件（文章による説明でも可）；用地境界線；その他提出すべき資料。受取専門：建築、構造、水、暖、電、智能化、経済",
        "ko": "총도 전문이 제출합니다. 내용: 부지 전체 배치(부지 경계선 및 경계선 내 기존 건축물, 구조물, 도로, 녹화 등의 명칭과 배치 상황); 간략한 도시 인프라 조건(문자 설명 가능); 용지 경계선; 기타 제출해야 할 자료. 수신 전문: 건축, 구조, 급배수, 난방, 전기, 지능화, 경제",
        "fr": "Présenté par la discipline d'aménagement du site. Contenu : plan général du site (limites du terrain et noms et implantation des bâtiments, ouvrages, voies, espaces verts existants dans le périmètre) ; brèves conditions des réseaux urbains (peuvent être décrites par écrit) ; limite d'utilisation du terrain ; autres documents à fournir. Disciplines destinataires : architecture, structure, plomberie, CVC, électricité, systèmes intelligents, économie",
        "de": "Vorgelegt von der Disziplin Standortplanung. Inhalt: Gesamtplan des Geländes (Grundstücksgrenze sowie Namen und Anordnung der vorhandenen Gebäude, Bauwerke, Straßen, Grünanlagen usw. innerhalb der Grenze); kurze Angaben zu den kommunalen Anschlussbedingungen (können textlich erfolgen); Nutzungsgrenze des Grundstücks; sonstige vorzulegende Unterlagen. Empfangende Disziplinen: Architektur, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de urbanización. Contenido: plano general del terreno (límite del terreno y nombres y disposición de los edificios, construcciones, vías, zonas verdes existentes dentro del límite); condiciones municipales resumidas (pueden expresarse por escrito); límite de uso del suelo; demás documentación a aportar. Especialidades receptoras: arquitectura, estructuras, fontanería, climatización, electricidad, sistemas inteligentes, economía",
        "pt": "Apresentado pela especialidade de urbanismo. Conteúdo: planta geral do terreno (limite do terreno e nomes e disposição das edificações, obras, vias, áreas verdes existentes dentro do limite); condições municipais resumidas (podem ser descritas por escrito); limite de uso do solo; demais documentos a fornecer. Especialidades receptoras: arquitetura, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes, economia",
        "ru": "Представляется специальностью генплана. Содержание: генеральный план участка (границы участка и наименования и расположение существующих зданий, сооружений, дорог, зелёных насаждений и т.п. в пределах границ); краткие условия городских сетей (допускается текстовое описание); границы землеотвода; прочие материалы, подлежащие представлению. Принимающие специальности: архитектура, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы, экономика",
        "ar": "يقدمه تخصص تخطيط الموقع. المحتوى: المخطط العام للموقع (حدود الموقع وأسماء وترتيب المباني والمنشآت والطرق والمساحات الخضراء القائمة داخل الحدود)؛ والشروط البلدية المختصرة (يمكن وصفها كتابياً)؛ وحدود استخدام الأرض؛ وغيرها من المواد التي يلزم تقديمها. التخصصات المستلمة: العمارة، الإنشاءات، السباكة، التدفئة والتهوية، الكهرباء، الأنظمة الذكية، الاقتصاد",
        "it": "Presentato dalla disciplina di pianificazione del sito. Contenuto: planimetria generale del sito (limite dell'area e nomi e disposizione di edifici, opere, strade, aree verdi esistenti entro il limite); brevi condizioni dei servizi urbani (possono essere fornite in forma testuale); limite di utilizzo del terreno; altri documenti da fornire. Discipline destinatarie: architettura, strutture, impianti idrici, HVAC, elettrici, sistemi intelligenti, economia",
        "vi": "Do chuyên ngành tổng mặt bằng đưa ra. Nội dung: tổng mặt bằng khu đất (chỉ giới khu đất và tên gọi, bố trí của nhà, công trình, đường, cây xanh hiện có trong chỉ giới); điều kiện hạ tầng đô thị tóm tắt (có thể diễn đạt bằng văn bản); chỉ giới sử dụng đất; các tài liệu khác cần nêu. Các chuyên ngành nhận: kiến trúc, kết cấu, cấp thoát nước, điều hòa không khí, điện, hệ thống thông minh, kinh tế",
        "th": "เสนอโดยสาขาการวางผังบริเวณ เนื้อหา: ผังบริเวณโดยรวม (แนวเขตที่ดิน และชื่อและการจัดวางของอาคาร สิ่งก่อสร้าง ถนน ภูมิทัศน์ที่มีอยู่ภายในแนวเขต) เงื่อนไขสาธารณูปโภคโดยย่อ (อาจอธิบายเป็นข้อความได้) แนวเขตการใช้ที่ดิน เอกสารอื่นที่ต้องเสนอ สาขาที่รับ: สถาปัตยกรรม โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ เศรษฐศาสตร์",
    },
    "项目经理组织、设计主持人主持、各专业负责人参加（第二阶段开始日）：第一阶段情况总结；研究和讨论存在的问题，介绍各专业反馈意见的协调情况和结果；发放第一版建筑作业图": {
        "en": "Organized by the project manager, presided over by the design director, attended by the person in charge of each discipline (starting day of the second stage): summary of the first-stage situation; study and discuss existing problems, and introduce the coordination status and results of the feedback opinions of each discipline; issue the first-version architectural working drawings",
        "ja": "プロジェクトマネージャーが組織し、設計主催者が司会し、各専門責任者が参加する（第二段階開始日）：第一段階の状況を総括する；存在する問題を研究・討論し、各専門のフィードバック意見の調整状況と結果を紹介する；第一版建築作業図を配布する",
        "ko": "프로젝트 매니저가 조직하고 설계 주재자가 사회하며 각 전문 책임자가 참석한다(제2단계 시작일)：제1단계 상황을 총결한다；존재하는 문제를 연구·토론하고 각 전문의 피드백 의견 조정 상황과 결과를 소개한다；제1판 건축 작업도를 배포한다",
        "fr": "Organisée par le chef de projet, présidée par le directeur de conception, avec la participation du responsable de chaque spécialité (jour de début de la deuxième étape) : synthèse de la situation de la première étape ; étude et discussion des problèmes existants, présentation de l'état et des résultats de la coordination des avis de chaque spécialité ; distribution de la première version des plans de travail d'architecture",
        "de": "Organisiert durch den Projektmanager, geleitet durch den Entwurfsleiter, unter Teilnahme des Verantwortlichen jeder Fachrichtung (Starttag der zweiten Phase): Zusammenfassung der Situation der ersten Phase; Untersuchung und Diskussion bestehender Probleme, Vorstellung des Abstimmungsstands und der Ergebnisse der Rückmeldungen jeder Fachrichtung; Ausgabe der ersten Fassung der Architektur-Arbeitspläne",
        "es": "Organizada por el gerente del proyecto, presidida por el director de diseño, con la participación del responsable de cada especialidad (día de inicio de la segunda etapa): resumen de la situación de la primera etapa; estudio y discusión de los problemas existentes, presentación del estado y los resultados de la coordinación de las opiniones de cada especialidad; distribución de la primera versión de los planos de trabajo de arquitectura",
        "pt": "Organizada pelo gerente do projeto, presidida pelo diretor de projeto, com a participação do responsável de cada especialidade (dia de início da segunda etapa): resumo da situação da primeira etapa; estudo e discussão dos problemas existentes, apresentação do estado e dos resultados da coordenação das opiniões de cada especialidade; distribuição da primeira versão dos desenhos de trabalho de arquitetura",
        "ru": "Организует руководитель проекта, ведет главный инженер проекта, участвуют руководители каждой специальности (в день начала второго этапа): обобщение итогов первого этапа; изучение и обсуждение существующих проблем, представление состояния и результатов координации замечаний каждой специальности; выдача первой версии архитектурных рабочих чертежей",
        "ar": "ينظمه مدير المشروع، ويترأسه مدير التصميم، ويحضره مسؤول كل تخصص (يوم بدء المرحلة الثانية): تلخيص حالة المرحلة الأولى؛ دراسة ومناقشة المشكلات القائمة وعرض حالة ونتائج التنسيق لآراء كل تخصص؛ توزيع النسخة الأولى من رسومات العمل المعمارية",
        "it": "Organizzata dal project manager, presieduta dal direttore di progettazione, con la partecipazione del responsabile di ciascuna disciplina (giorno di inizio della seconda fase): sintesi della situazione della prima fase; studio e discussione dei problemi esistenti, presentazione dello stato e dei risultati del coordinamento delle osservazioni di ciascuna disciplina; emissione della prima versione degli elaborati operativi architettonici",
        "vi": "Do giám đốc dự án tổ chức, chủ trì thiết kế chủ trì, các trưởng chuyên ngành tham dự (ngày bắt đầu giai đoạn thứ hai): tổng kết tình hình giai đoạn thứ nhất; nghiên cứu và thảo luận các vấn đề tồn tại, giới thiệu tình hình và kết quả điều phối ý kiến phản hồi của từng chuyên ngành; phát hành bản vẽ tác nghiệp kiến trúc lần thứ nhất",
        "th": "จัดโดยผู้จัดการโครงการ เป็นประธานโดยผู้อำนวยการออกแบบ ผู้รับผิดชอบแต่ละสาขาเข้าร่วม (วันเริ่มขั้นตอนที่สอง): สรุปสถานการณ์ของขั้นตอนที่หนึ่ง ศึกษาและอภิปรายปัญหาที่มีอยู่ แนะนำสถานะและผลการประสานความคิดเห็นของแต่ละสาขา แจกจ่ายแบบปฏิบัติงานสถาปัตยกรรมฉบับแรก",
    },
    "建筑专业提出，作为各专业初步设计工作基础用图。内容（在第一阶段基础上）：各层建筑平面图、主要立面图及剖面图（含初步防火分区示意、初步人防平面及口部布置、承重结构位置轴线与尺寸、机电能源站房及分散机房位置、主要建筑构配件、主要建筑设备位置、房间名称与设计标高、二次装修区域范围、主要管道井配电间弱电间位置、屋顶设备与屋面绿化说明等）；文字说明（楼地面墙体屋面及地下室底板初步构造做法、不同功能区域建筑面积和（或）设计使用人数）。接受专业：总图、结构、水、暖、电、智能化": {
        "en": "Proposed by the architectural discipline as the base drawings for the preliminary design work of all disciplines. Contents (based on the first stage): architectural plans of each floor, main elevations and sections (including preliminary fire compartmentation diagrams, preliminary civil defence plans and entrance layouts, axes and dimensions of load-bearing structure positions, locations of MEP energy plant rooms and scattered equipment rooms, main architectural components, locations of main building equipment, room names and design floor elevations, scope of secondary decoration areas, locations of main riser shafts, electrical distribution rooms and weak-current rooms, rooftop equipment and roof greening notes, etc.); textual notes (preliminary construction practices for floors, walls, roofs and basement slabs, and gross floor areas and/or design occupant numbers of different functional zones). Receiving disciplines: site planning, structural, plumbing, HVAC, electrical, intelligent systems",
        "ja": "建築専門が提出し、各専門の基本設計作業の基図に供する。内容（第1段階に基づく）：各階建築平面図、主要立面図・断面図（初步防火区画図、初步人防平面および口部配置、耐力構造位置の軸線と寸法、設備エネルギー供給ステーションおよび分散機械室の位置、主要建築部材・設備の位置、室名と設計標高、二次仕上げ区域範囲、主要パイプシャフト・配電室・弱電室の位置、屋上設備と屋上緑化の説明等を含む）；文字説明（床・壁・屋根・地下階床版の初步的構造仕様、異なる機能区域の建築面積および（または）設計利用人数）。受取専門：総図、構造、水、暖、電、智能化",
        "ko": "건축 전문이 제출하며 각 전문의 기본설계 작업 기초도면으로 사용됩니다. 내용(1단계에 기초): 각 층 건축 평면도, 주요 입면도 및 단면도(초기 방화구획도, 초기 인방(人防) 평면 및 출입구 배치, 내력 구조 위치 축선과 치수, 기계·전기 에너지 스테이션 및 분산 기계실 위치, 주요 건축 부재·설비 위치, 실명과 설계 표고, 2차 장식 구역 범위, 주요 파이프 샤프트·배전실·약전실 위치, 옥상 설비 및 옥상 녹화 설명 등 포함); 문자 설명(바닥·벽·지붕·지하층 바닥판 초기 구성 공법, 기능 구역별 건축 면적 및(또는) 설계 사용 인원). 수신 전문: 총도, 구조, 급배수, 난방, 전기, 지능화",
        "fr": "Présenté par la discipline architecture, servant de plans de base pour les travaux de projet d'esquisse de toutes les disciplines. Contenu (sur la base de la première phase) : plans d'architecture de chaque niveau, élévations et coupes principales (y compris schémas préliminaires de compartimentage coupe-feu, plan préliminaire de la défense civile et dispositions des entrées, axes et dimensions des emplacements de la structure porteuse, emplacements des locaux techniques énergétiques et des locaux techniques dispersés, principaux composants architecturaux, emplacements des principaux équipements du bâtiment, noms des pièces et niveaux finis de conception, étendue des zones de second œuvre, emplacements des principales gaines techniques, des locaux électriques et des locaux courants faibles, notes sur les équipements de toiture et la végétalisation, etc.) ; notes écrites (pratiques constructives préliminaires des planchers, murs, toitures et dalles de sous-sol, surfaces hors œuvre et/ou effectifs de conception des différentes zones fonctionnelles). Disciplines destinataires : aménagement du site, structure, plomberie, CVC, électricité, systèmes intelligents",
        "de": "Vorgelegt von der Architekturdisziplin als Arbeitsgrundlage für die Vorplanung aller Fachdisziplinen. Inhalt (aufbauend auf der ersten Stufe): Architekturpläne aller Geschosse, Hauptansichten und -schnitte (einschließlich vorläufiger Brandabschnittsdarstellungen, vorläufiger Zivilschutzpläne und Zugangsanordnungen, Achsen und Abmessungen der Positionen der tragenden Struktur, Standorte der TGA-Energiezentralen und dezentraler Technikräume, wesentlicher Bauteile, Standorte der wesentlichen Gebäudeausrüstung, Raumbezeichnungen und Fertigfußbodenhöhen, Umfang der Sekundärausbauzonen, Standorte der Hauptschächte, Elektro- und Schwachstromräume, Hinweise zu Dachausrüstung und Dachbegrünung usw.); textliche Erläuterungen (vorläufige Konstruktionsarten für Böden, Wände, Dächer und Untergeschossplatten sowie Bruttoflächen und/oder Bemessungsbelegungszahlen der verschiedenen Funktionsbereiche). Empfangende Disziplinen: Standortplanung, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation",
        "es": "Presentado por la especialidad de arquitectura, como planos base para los trabajos de diseño preliminar de todas las especialidades. Contenido (sobre la base de la primera fase): plantas de arquitectura de cada nivel, alzados y secciones principales (incluidos esquemas preliminares de compartimentación contra incendios, plano preliminar de defensa civil y disposición de accesos, ejes y dimensiones de la estructura portante, ubicación de salas de energía y locales técnicos dispersos, principales componentes arquitectónicos, ubicación de los principales equipos del edificio, nombres de locales y cotas de proyecto, alcance de las zonas de acabado secundario, ubicación de patinillos principales, cuartos eléctricos y cuartos de corrientes débiles, notas sobre equipos de cubierta y cubiertas verdes, etc.); notas escritas (prácticas constructivas preliminares de suelos, muros, cubiertas y losas de sótano, superficies construidas y/o aforos de diseño de las distintas zonas funcionales). Especialidades receptoras: urbanización, estructuras, fontanería, climatización, electricidad, sistemas inteligentes",
        "pt": "Apresentado pela especialidade de arquitetura, como desenhos base para os trabalhos de projeto preliminar de todas as especialidades. Conteúdo (com base na primeira fase): plantas de arquitetura de cada pavimento, fachadas e cortes principais (incluindo esquemas preliminares de compartimentação contra incêndio, planta preliminar de defesa civil e disposição de acessos, eixos e dimensões da estrutura portante, localização das centrais de energia e salas técnicas dispersas, principais componentes arquitetônicos, localização dos principais equipamentos do edifício, nomes dos ambientes e cotas de projeto, extensão das áreas de acabamento secundário, localização dos principais shaft, salas elétricas e salas de corrente fraca, notas sobre equipamentos de cobertura e cobertura verde, etc.); notas escritas (práticas construtivas preliminares de pisos, paredes, coberturas e lajes de subsolo, áreas construídas e/ou lotações de projeto das diferentes zonas funcionais). Especialidades receptoras: urbanismo, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes",
        "ru": "Представляется архитектурной специальностью в качестве базовых чертежей для работ по разработке стадии «проект» всеми специальностями. Содержание (на основе первого этапа): архитектурные планы всех этажей, основные фасады и разрезы (включая предварительные схемы противопожарных отсеков, предварительный план гражданской обороны и расположение входов, оси и размеры расположения несущих конструкций, расположение энергетических станций и отдельных технических помещений, основные архитектурные элементы, расположение основного инженерного оборудования здания, наименования помещений и проектные отметки, границы зон вторичной отделки, расположение основных инженерных шахт, электрощитовых и слаботочных помещений, пояснения по оборудованию кровли и озеленению крыши и т.п.); текстовые пояснения (предварительные конструктивные решения полов, стен, кровли и плит подвала, площади и/или расчётная численность пользователей различных функциональных зон). Принимающие специальности: генплан, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы",
        "ar": "يقدمه تخصص العمارة كمخططات أساسية لأعمال التصميم الأولي لجميع التخصصات. المحتوى (استناداً إلى المرحلة الأولى): مخططات أفقية معمارية لكل طابق، وواجهات وقطاعات رئيسية (بما في ذلك مخططات أولية لتقسيم الحماية من الحرائق، ومخطط أولي للدفاع المدني وترتيب المداخل، ومحاور وأبعاد مواقع الإنشاء الحامل، ومواقع محطات الطاقة و الغرف الفنية الموزعة، والمكونات المعمارية الرئيسية، ومواقع المعدات الرئيسية للمبنى، وأسماء الغرف ومناسيب التصميم، ونطاق مناطق التشطيبات الثانوية، ومواقع الأعمدة الرئيسية وغرف التوزيع الكهربائي وغرف التيارات الضعيفة، وملاحظات معدات الأسطح وتنسيق الأسطح الخضراء وغيرها)؛ وملاحظات نصية (أساليب إنشائية أولية للأرضيات والجدران والأسطح وألواح الطابق السفلي، ومساحات البناء و/أو أعداد المستخدمين التصميمية لمختلف المناطق الوظيفية). التخصصات المستلمة: تخطيط الموقع، الإنشاءات، السباكة، التدفئة والتهوية، الكهرباء، الأنظمة الذكية",
        "it": "Presentato dalla disciplina architettonica come elaborati di base per il lavoro di progettazione preliminare di tutte le discipline. Contenuto (sulla base della prima fase): piante architettoniche di ogni piano, prospetti e sezioni principali (inclusi schemi preliminari di compartimentazione antincendio, pianta preliminare di protezione civile e disposizione degli accessi, assi e dimensioni delle posizioni della struttura portante, posizioni delle centrali impiantistiche e dei locali tecnici distribuiti, principali componenti architettonici, posizioni delle principali attrezzature dell'edificio, nomi degli ambienti e quote di progetto, estensione delle aree di finitura secondaria, posizioni dei principali vani tecnici, locali elettrici e locali debole corrente, note su attrezzature di copertura e verde pensile, ecc.); note testuali (pratiche costruttive preliminari per pavimenti, pareti, coperture e soletta di piano interrato, superfici lorde e/o occupanti di progetto delle diverse zone funzionali). Discipline destinatarie: pianificazione del sito, strutture, impianti idrici, HVAC, elettrici, sistemi intelligenti",
        "vi": "Do chuyên ngành kiến trúc đưa ra, dùng làm bản vẽ cơ sở cho công tác thiết kế cơ sở của các chuyên ngành. Nội dung (trên cơ sở giai đoạn 1): mặt bằng kiến trúc từng tầng, mặt đứng và mặt cắt chính (bao gồm sơ đồ phân vùng chống cháy sơ bộ, mặt bằng phòng thủ dân sự sơ bộ và bố trí cửa, trục và kích thước vị trí kết cấu chịu lực, vị trí trạm năng lượng cơ điện và phòng máy phân tán, cấu kiện kiến trúc chính, vị trí thiết bị chính của công trình, tên phòng và cao độ thiết kế, phạm vi khu vực trang trí thứ cấp, vị trí giếng kỹ thuật chính, phòng điện và phòng yếu điện, ghi chú thiết bị mái và xanh hóa mái, v.v.); ghi chú văn bản (cấu tạo sơ bộ sàn, tường, mái và bản đáy tầng hầm, diện tích sàn và/hoặc số người sử dụng thiết kế của các khu vực chức năng khác nhau). Các chuyên ngành nhận: tổng mặt bằng, kết cấu, cấp thoát nước, điều hòa không khí, điện, hệ thống thông minh",
        "th": "เสนอโดยสาขาสถาปัตยกรรม ใช้เป็นแบบพื้นฐานสำหรับงานออกแบบขั้นต้นของทุกสาขา เนื้อหา (ต่อจากระยะที่ 1): แบบพื้นสถาปัตยกรรมแต่ละชั้น แบบด้านและแบบตัดหลัก (รวมถึงผังการแบ่งเขตอัคคีภัยเบื้องต้น ผังป้องกันภัยทางอากาศเบื้องต้นและการจัดตำแหน่งทางเข้า แกนและขนาดตำแหน่งโครงสร้างรับน้ำหนัก ตำแหน่งห้องพลังงานงานระบบและห้องเครื่องกระจาย องค์ประกอบสถาปัตยกรรมหลัก ตำแหน่งอุปกรณ์หลักของอาคาร ชื่อห้องและระดับพื้นสำเร็จรูป ขอบเขตพื้นที่ตกแต่งขั้นที่สอง ตำแหน่งปล่องท่อหลัก ห้องไฟฟ้า ห้องระบบไฟฟ้าอ่อน หมายเหตุอุปกรณ์บนหลังคาและหลังคาเขียว ฯลฯ) หมายเหตุเป็นข้อความ (วิธีก่อสร้างเบื้องต้นของพื้น ผนัง หลังคา และแผ่นพื้นชั้นใต้ดิน พื้นที่อาคารและ/หรือจำนวนผู้ใช้ตามการออกแบบของพื้นที่ใช้งานต่าง ๆ) สาขาที่รับ: การวางผังบริเวณ โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ",
    },
    "水提暖：特殊房间温湿度参数和通风要求（如水箱间温度要求、气体消防房间名称等）；暖提水：不采暖房间名称；电提暖：特殊房间温湿度参数和通风要求、电气发热量较大的房间设备发热量（或估算电气安装容量，如变配电室、大型计算机主机房）；智能化提暖：各弱电机房、安防控制室温湿度参数和通风要求；智能化提电：各弱电机房、安防控制室、弱电间用电量": {
        "en": "Plumbing to HVAC: temperature and humidity parameters and ventilation requirements of special rooms (such as temperature requirements of water tank rooms, names of gas fire-protection rooms, etc.); HVAC to plumbing: names of unheated rooms; electrical to HVAC: temperature and humidity parameters and ventilation requirements of special rooms, heat gain of equipment in rooms with large electrical heat release (or estimated electrical installation capacity, such as substations and large computer mainframe rooms); intelligent systems to HVAC: temperature and humidity parameters and ventilation requirements of weak-current equipment rooms and security control rooms; intelligent systems to electrical: power consumption of weak-current equipment rooms, security control rooms and weak-current rooms",
        "ja": "水→暖：特殊室の温湿度パラメータと換気要求（水槽室の温度要求、ガス系消防室の名称等）；暖→水：非暖房室の名称；電→暖：特殊室の温湿度パラメータと換気要求、電気発熱量が大きい室の機器発熱量（または概算電気設備容量。変配電室、大型コンピュータ室など）；智能化→暖：各弱電機械室・安防控制室の温湿度パラメータと換気要求；智能化→電：各弱電機械室・安防控制室・弱電間の電力消費量",
        "ko": "급배수→난방: 특수실의 온습도 파라미터와 환기 요구(예: 수조실 온도 요구, 가스 소방실 명칭 등); 난방→급배수: 비난방실 명칭; 전기→난방: 특수실 온습도 파라미터와 환기 요구, 전기 발열량이 큰 실의 설비 발열량(또는 추정 전기 설비 용량, 예: 변배전실, 대형 컴퓨터 주기기실); 지능화→난방: 각 약전 기계실·보안 통제실의 온습도 파라미터와 환기 요구; 지능화→전기: 각 약전 기계실·보안 통제실·약전실의 전력 소비량",
        "fr": "Plomberie vers CVC : paramètres de température et d'humidité et exigences de ventilation des locaux spéciaux (tels que les exigences de température des locaux de réservoirs d'eau, les noms des locaux de protection incendie à gaz, etc.) ; CVC vers plomberie : noms des locaux non chauffés ; électricité vers CVC : paramètres de température et d'humidité et exigences de ventilation des locaux spéciaux, dégagement thermique des équipements des locaux à forte charge électrique (ou capacité électrique installée estimée, comme les postes de transformation et les grandes salles d'ordinateurs centraux) ; systèmes intelligents vers CVC : paramètres de température et d'humidité et exigences de ventilation des locaux techniques courants faibles et des salles de contrôle de sécurité ; systèmes intelligents vers électricité : consommation électrique des locaux techniques courants faibles, des salles de contrôle de sécurité et des gaines courants faibles",
        "de": "Sanitär an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter sowie Lüftungsanforderungen besonderer Räume (z. B. Temperaturanforderungen von Wasserbehälterräumen, Bezeichnungen von Gaslöschräumen usw.); Heizung-Lüftung an Sanitär: Bezeichnungen unbeheizter Räume; Elektro an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter sowie Lüftungsanforderungen besonderer Räume, Wärmeabgabe von Geräten in Räumen mit hoher elektrischer Wärmeentwicklung (oder geschätzte elektrische Anschlussleistung, z. B. Umspannwerke und große Rechnerräume); Gebäudeautomation an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter sowie Lüftungsanforderungen von Schwachstrom-Räumen und Sicherheitskontrollräumen; Gebäudeautomation an Elektro: Stromverbrauch von Schwachstrom-Räumen, Sicherheitskontrollräumen und Schwachstrom-Schächten",
        "es": "Fontanería a climatización: parámetros de temperatura y humedad y requisitos de ventilación de locales especiales (como requisitos de temperatura de cuartos de tanques de agua, nombres de locales de protección contra incendios por gas, etc.); climatización a fontanería: nombres de locales sin calefacción; electricidad a climatización: parámetros de temperatura y humedad y requisitos de ventilación de locales especiales, carga térmica de los equipos de locales con gran disipación eléctrica (o capacidad eléctrica instalada estimada, como subestaciones y grandes salas de computadoras); sistemas inteligentes a climatización: parámetros de temperatura y humedad y requisitos de ventilación de los cuartos de corrientes débiles y salas de control de seguridad; sistemas inteligentes a electricidad: consumo eléctrico de los cuartos de corrientes débiles, salas de control de seguridad y patinillos de corrientes débiles",
        "pt": "Hidráulica para climatização: parâmetros de temperatura e umidade e requisitos de ventilação de ambientes especiais (como requisitos de temperatura de salas de reservatórios de água, nomes de ambientes de proteção contra incêndio por gás, etc.); climatização para hidráulica: nomes de ambientes sem aquecimento; elétrica para climatização: parâmetros de temperatura e umidade e requisitos de ventilação de ambientes especiais, carga térmica dos equipamentos de ambientes com grande dissipação elétrica (ou capacidade elétrica instalada estimada, como subestações e grandes salas de computadores); sistemas inteligentes para climatização: parâmetros de temperatura e umidade e requisitos de ventilação das salas de corrente fraca e salas de controle de segurança; sistemas inteligentes para elétrica: consumo elétrico das salas de corrente fraca, salas de controle de segurança e shaft de corrente fraca",
        "ru": "Водоснабжение — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции специальных помещений (например, требования к температуре помещений водонапорных баков, наименования помещений газового пожаротушения и т.п.); отопление и вентиляция — водоснабжению: наименования неотапливаемых помещений; электроснабжение — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции специальных помещений, тепловыделение оборудования в помещениях с большим тепловыделением от электрооборудования (или расчётная установленная мощность, например трансформаторные подстанции, крупные машинные залы); интеллектуальные системы — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции слаботочных аппаратных и помещений охраны; интеллектуальные системы — электроснабжению: потребляемая мощность слаботочных аппаратных, помещений охраны и слаботочных помещений",
        "ar": "السباكة إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية للغرف الخاصة (مثل متطلبات درجة حرارة غرف خزانات المياه، وأسماء غرف الإطفاء بالغاز وغيرها)؛ والتدفئة والتهوية إلى السباكة: أسماء الغرف غير المدفأة؛ والكهرباء إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية للغرف الخاصة، والحرارة المنبعثة من معدات الغرف ذات الإطلاق الكهربائي الكبير (أو السعة الكهربائية المركبة التقديرية، مثل محطات التحويل وغرف الحواسيب الرئيسية الكبيرة)؛ والأنظمة الذكية إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية لغرف التيارات الضعيفة وغرف التحكم الأمنية؛ والأنظمة الذكية إلى الكهرباء: استهلاك الكهرباء لغرف التيارات الضعيفة وغرف التحكم الأمنية وأعمدة التيارات الضعيفة",
        "it": "Impianti idrici a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali speciali (come i requisiti di temperatura dei locali serbatoi acqua, i nomi dei locali di protezione antincendio a gas, ecc.); HVAC a impianti idrici: nomi dei locali non riscaldati; impianti elettrici a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali speciali, carico termico delle apparecchiature dei locali con grande dissipazione elettrica (o potenza elettrica installata stimata, come cabine elettriche e grandi sale computer); sistemi intelligenti a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali debole corrente e delle sale controllo sicurezza; sistemi intelligenti a elettrici: consumo elettrico dei locali debole corrente, delle sale controllo sicurezza e dei vani debole corrente",
        "vi": "Cấp thoát nước gửi điều hòa không khí: thông số nhiệt độ độ ẩm và yêu cầu thông gió của phòng đặc biệt (như yêu cầu nhiệt độ phòng bể nước, tên phòng chữa cháy bằng khí, v.v.); điều hòa gửi cấp thoát nước: tên phòng không sưởi; điện gửi điều hòa: thông số nhiệt độ độ ẩm và yêu cầu thông gió của phòng đặc biệt, lượng nhiệt tỏa ra của thiết bị trong phòng có lượng tỏa nhiệt điện lớn (hoặc công suất lắp đặt điện ước tính, như trạm biến áp, phòng máy tính lớn); hệ thống thông minh gửi điều hòa: thông số nhiệt độ độ ẩm và yêu cầu thông gió của các phòng máy yếu điện và phòng điều khiển an ninh; hệ thống thông minh gửi điện: lượng điện tiêu thụ của các phòng máy yếu điện, phòng điều khiển an ninh và phòng yếu điện",
        "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่มีการทำความร้อน ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ปริมาณความร้อนจากอุปกรณ์ในห้องที่มีการปลดปล่อยความร้อนทางไฟฟ้าสูง (หรือกำลังไฟฟ้าติดตั้งโดยประมาณ เช่น ห้องสถานีย่อย ห้องคอมพิวเตอร์เมนเฟรมขนาดใหญ่) ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องอุปกรณ์ไฟฟ้าอ่อนและห้องควบคุมระบบรักษาความปลอดภัย ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณการใช้ไฟฟ้าของห้องอุปกรณ์ไฟฟ้าอ่อน ห้องควบคุมระบบรักษาความปลอดภัย และปล่องไฟฟ้าอ่อน",
    },
    "水提建筑：补充完善确认水泵房、水处理机房、热交换站、热水机房、水箱间、报警阀室等机房面积尺寸净高要求，集中管道井位置尺寸，水池尺寸与平面定位、地下室集水坑位置尺寸，影响室内布置和外立面的设备（消火栓、冷却塔等）初步位置，建筑最底层垫层厚度要求；水提暖：由动力专业提供热源时提出估算总热负荷和二次水温度参数（热媒为蒸汽时提出凝结水回水方式）；暖提建筑：补充完善确认冷冻机房、水泵房、锅炉房、热交换站、燃气调压站（室）等机房面积尺寸净高，各层分散空调通风机房位置面积，集中管道井位置尺寸，锅炉房等泄爆井（泄爆面）初步位置面积；暖提水：主要用水点房间名称和估算用水量、主要排水点房间名称和排水性质、冷却水总流量进出水温要求和拟定冷水机组台数；暖提电：需防雷接地和防静电的设备名称（油气锅炉房、烟囱、燃气放散管、可燃气体管道等）；电提建筑：补充完善确认变配电所、发电机房等机房面积尺寸净高，各层弱电机房配电小室位置面积，主要强电竖井位置尺寸，楼板厚度要求，发电机冷却进风与排风竖井位置面积；电提水：特殊电气用房消防要求；智能化提建筑：补充完善确认进线间、网络机房、电话交换机房、有线电视机房、手机信号覆盖机房、安防控制室、卫星电视等机房面积尺寸净高，弱电竖井大型会议室控制室位置尺寸，楼板厚度要求": {
        "en": "Plumbing to architectural: supplement, improve and confirm the area, dimensions and clear height requirements of pump rooms, water treatment rooms, heat exchange stations, hot water rooms, water tank rooms, alarm valve rooms and other plant rooms, the position and dimensions of central riser shafts, the dimensions and plan location of water tanks, the position and dimensions of basement sump pits, the preliminary positions of equipment affecting interior layout and elevations (fire hydrants, cooling towers, etc.), and the thickness requirement of the lowest floor cushion layer; plumbing to HVAC: when the heat source is provided by the utility discipline, raise the estimated total heat load and the secondary water temperature parameters (when the heating medium is steam, raise the condensate return method); HVAC to architectural: supplement, improve and confirm the area, dimensions and clear height of refrigeration rooms, pump rooms, boiler rooms, heat exchange stations, gas pressure regulating stations (rooms), etc., the positions and areas of decentralized air-conditioning and ventilation rooms on each floor, the position and dimensions of central riser shafts, and the preliminary position and area of relief shafts (relief surfaces) of boiler rooms, etc.; HVAC to plumbing: the room names and estimated water consumption of main water use points, the room names and drainage nature of main drainage points, the total cooling water flow rate, inlet/outlet water temperature requirements and the proposed number of chiller units; HVAC to electrical: the names of equipment requiring lightning protection, earthing and anti-static measures (oil and gas boiler rooms, chimneys, gas vent pipes, combustible gas pipelines, etc.); electrical to architectural: supplement, improve and confirm the area, dimensions and clear height of substations, generator rooms and other plant rooms, the positions and areas of weak-current rooms and distribution cubicles on each floor, the position and dimensions of main power riser shafts, the floor slab thickness requirement, and the positions and areas of generator cooling air intake and exhaust shafts; electrical to plumbing: fire protection requirements for special electrical rooms; intelligent systems to architectural: supplement, improve and confirm the area, dimensions and clear height of incoming line rooms, network rooms, telephone exchange rooms, cable TV rooms, mobile signal coverage rooms, security control rooms, satellite TV rooms, etc., the position and dimensions of weak-current riser shafts, large meeting room control rooms, and the floor slab thickness requirement",
        "ja": "水提建築：ポンプ室・水処理室・熱交換ステーション・温水室・水槽室・警報弁室等の機械室面積・寸法・天井高要求を補完・改善・確認し、集中パイプシャフトの位置寸法、水槽の寸法と平面位置、地下集水桝の位置寸法、室内配置と外立面に影響する設備（消火栓・冷却塔等）の初期位置、建築最下層の捨てコンクリート厚さ要求を提出する；水提暖：動力専門から熱源供給時、推定総熱負荷と二次水温度パラメータを提出する（熱媒が蒸気の場合、復水の還水方式を提出する）；暖提建築：冷凍室・ポンプ室・ボイラー室・熱交換ステーション・ガス調圧ステーション（室）等の機械室面積・寸法・天井高を補完・改善・確認し、各階の分散空調換気室の位置面積、集中パイプシャフトの位置寸法、ボイラー室等の爆発逃がしシャフト（逃がし面）の初期位置面積を提出する；暖提水：主要用水点の室名と推定用水量、主要排水点の室名と排水性質、冷却水総流量・出入口水温要求と予定冷凍機台数を提出する；暖提電：防雷接地・防静電を要する設備名（油ガスボイラー室・煙突・ガス放散管・可燃性ガス配管等）を提出する；電提建築：変電所・発電機室等の機械室面積・寸法・天井高を補完・改善・確認し、各階の弱電室・配電小室の位置面積、主要強電シャフトの位置寸法、床スラブ厚さ要求、発電機冷却の給気・排気シャフトの位置面積を提出する；電提水：特殊電気室の消防要求を提出する；智能化提建築：引込線室・ネットワーク室・電話交換室・有線テレビ室・携帯信号カバー室・安防控制室・衛星テレビ等の機械室面積・寸法・天井高を補完・改善・確認し、弱電シャフト・大型会議室制御室の位置寸法、床スラブ厚さ要求を提出する",
        "ko": "급배수가 건축에: 펌프실·수처리실·열교환 스테이션·온수실·물탱크실·경보밸브실 등 기계실의 면적·치수·천장고 요구를 보완·개선·확인하고, 중앙 파이프 샤프트 위치 치수, 물탱크 치수와 평면 위치, 지하 집수정 위치 치수, 실내 배치와 외관에 영향을 주는 설비(소화전·냉각탑 등)의 초기 위치, 건축 최하층 바닥 콘크리트 두께 요구를 제출합니다; 급배수가 난방에: 동력 전문이 열원을 공급할 때 추정 총 열부하와 2차수 온도 파라미터를 제출합니다(열매체가 증기일 때 응축수 회수 방식을 제출합니다); 난방이 건축에: 냉동실·펌프실·보일러실·열교환 스테이션·가스 압력조정소(실) 등 기계실의 면적·치수·천장고를 보완·개선·확인하고, 각 층 분산 공조 환기실의 위치 면적, 중앙 파이프 샤프트 위치 치수, 보일러실 등 폭발 방출 샤프트(방출면)의 초기 위치 면적을 제출합니다; 난방이 급배수에: 주요 용수점 실명과 추정 용수량, 주요 배수점 실명과 배수 성질, 냉각수 총 유량·입출구 수온 요구와 예정 냉동기 대수를 제출합니다; 난방이 전기에: 방뢰 접지·정전기 방지가 필요한 설비명(유가스 보일러실·굴뚝·가스 방산관·가연성 가스 배관 등)을 제출합니다; 전기가 건축에: 변전소·발전기실 등 기계실의 면적·치수·천장고를 보완·개선·확인하고, 각 층 약전실·배전 소실의 위치 면적, 주요 강전 샤프트 위치 치수, 슬래브 두께 요구, 발전기 냉각 흡기·배기 샤프트 위치 면적을 제출합니다; 전기가 급배수에: 특수 전기실의 소방 요구를 제출합니다; 지능화가 건축에: 인입선실·네트워크실·전화교환실·유선TV실·휴대폰 신호 커버실·안보 통제실·위성TV 등 기계실의 면적·치수·천장고를 보완·개선·확인하고, 약전 샤프트·대형 회의실 제어실의 위치 치수, 슬래브 두께 요구를 제출합니다",
        "fr": "Plomberie à architecture : compléter, améliorer et confirmer les exigences de surface, de dimensions et de hauteur libre des locaux techniques tels que stations de pompage, locaux de traitement d'eau, postes d'échange thermique, locaux d'eau chaude, locaux de réservoirs d'eau, locaux de vannes d'alarme, les positions et dimensions des gaines techniques centrales, les dimensions et la localisation en plan des réservoirs d'eau, les positions et dimensions des puisards de sous-sol, les positions préliminaires des équipements influant sur l'aménagement intérieur et les façades (poteaux d'incendie, tours de refroidissement, etc.) et l'exigence d'épaisseur de la couche de forme du niveau le plus bas ; plomberie à CVC : lorsque la source de chaleur est fournie par la discipline énergétique, formuler la charge thermique totale estimée et les paramètres de température de l'eau secondaire (si le fluide caloporteur est de la vapeur, formuler le mode de retour des condensats) ; CVC à architecture : compléter, améliorer et confirmer la surface, les dimensions et la hauteur libre des locaux de réfrigération, stations de pompage, chaufferies, postes d'échange thermique, postes (locaux) de détente de gaz, etc., les positions et surfaces des locaux décentralisés de CVC et de ventilation à chaque niveau, les positions et dimensions des gaines techniques centrales, la position et la surface préliminaires des conduits (surfaces) de décompression des chaufferies, etc. ; CVC à plomberie : les noms des pièces des principaux points d'utilisation d'eau et l'estimation de la consommation d'eau, les noms des pièces des principaux points de drainage et la nature du drainage, le débit total de l'eau de refroidissement, les exigences de température d'entrée/sortie et le nombre proposé de groupes refroidisseurs ; CVC à électricité : les noms des équipements nécessitant une mise à la terre de protection contre la foudre et antistatique (chaufferies gazole/gaz, cheminées, cheminées de ventilation de gaz, canalisations de gaz combustible, etc.) ; électricité à architecture : compléter, améliorer et confirmer la surface, les dimensions et la hauteur libre des postes de transformation, groupes électrogènes et autres locaux techniques, les positions et surfaces des locaux courants faibles et des petites armoires de distribution à chaque niveau, les positions et dimensions des gaines électriques principales, l'exigence d'épaisseur des dalles, les positions et surfaces des conduits d'aspiration et d'évacuation du refroidissement des groupes électrogènes ; électricité à plomberie : les exigences de protection incendie des locaux électriques spéciaux ; systèmes intelligents à architecture : compléter, améliorer et confirmer la surface, les dimensions et la hauteur libre des locaux de pénétration, locaux réseau, locaux de centraux téléphoniques, locaux de télévision par câble, locaux de couverture de signaux mobiles, salles de contrôle de sécurité, locaux de télévision par satellite, etc., les positions et dimensions des gaines courants faibles et des salles de contrôle de grandes salles de réunion, et l'exigence d'épaisseur des dalles",
        "de": "Sanitär an Architektur: Ergänzung, Verbesserung und Bestätigung der Anforderungen an Fläche, Abmessungen und lichte Höhe von Pumpenräumen, Wasseraufbereitungsräumen, Wärmeübergabestationen, Warmwasserzentralen, Behälterräumen, Alarmventilräumen usw., der Position und Abmessungen zentraler Schächte, der Abmessungen und Grundrisslage von Wasserbehältern, der Position und Abmessungen von Kellerschächten, der vorläufigen Positionen von Anlagen, die Innenanordnung und Fassade beeinflussen (Hydranten, Kühltürme usw.), und der Anforderung an die Dicke der untersten Bodenschicht; Sanitär an Heizung-Lüftung: bei Wärmequelle durch die Energiedisziplin werden geschätzte Gesamtwärmelast und Sekundärwasser-Temperaturparameter angegeben (bei Dampf als Wärmeträger wird die Kondensatrückführungsart angegeben); Heizung-Lüftung an Architektur: Ergänzung, Verbesserung und Bestätigung der Fläche, Abmessungen und lichten Höhe von Kälteräumen, Pumpenräumen, Heizräumen, Wärmeübergabestationen, Gasdruckregelstationen (Räumen) usw., der Positionen und Flächen dezentraler Klima- und Lüftungsräume je Geschoss, der Position und Abmessungen zentraler Schächte, der vorläufigen Position und Fläche von Druckentlastungsschächten (Druckentlastungsflächen) der Heizräume usw.; Heizung-Lüftung an Sanitär: Raumnamen und geschätzter Wasserverbrauch der Hauptwasserentnahmepunkte, Raumnamen und Entwässerungsart der Hauptentwässerungspunkte, Gesamtdurchfluss des Kühlwassers, Vor- und Rücklauftemperaturanforderungen und die geplante Anzahl der Kältemaschinen; Heizung-Lüftung an Elektro: Namen der Anlagen mit Anforderungen an Blitzschutzerdung und Antistatik (Öl- und Gasheizräume, Schornsteine, Gasabblaseleitungen, brennbare Gasleitungen usw.); Elektro an Architektur: Ergänzung, Verbesserung und Bestätigung der Fläche, Abmessungen und lichten Höhe von Trafostationen, Generatorräumen usw., der Positionen und Flächen von Schwachstromräumen und Verteilerkabinen je Geschoss, der Position und Abmessungen der Hauptstromschächte, der Anforderung an die Geschossdeckenstärke und der Positionen und Flächen von Zuluft- und Abluftschächten der Generatorkühlung; Elektro an Sanitär: Brandschutzanforderungen für besondere Elektroräume; Gebäudeautomation an Architektur: Ergänzung, Verbesserung und Bestätigung der Fläche, Abmessungen und lichten Höhe von Einführungsräumen, Netzwerkräumen, Telefonvermittlungsräumen, Kabelfernsehräumen, Mobilfunksignale-Versorgungsräumen, Sicherheitskontrollräumen, Satellitenfernsehräumen usw., der Position und Abmessungen von Schwachstromschächten, Kontrollräumen großer Konferenzräume und der Anforderung an die Geschossdeckenstärke",
        "es": "Fontanería a arquitectura: completar, mejorar y confirmar los requisitos de superficie, dimensiones y altura libre de locales técnicos como salas de bombas, salas de tratamiento de agua, estaciones de intercambio térmico, salas de agua caliente, salas de depósitos de agua, salas de válvulas de alarma, etc., las posiciones y dimensiones de los patinillos técnicos centrales, las dimensiones y la ubicación en planta de los depósitos de agua, las posiciones y dimensiones de los sumideros de sótano, las posiciones preliminares de los equipos que afectan a la distribución interior y a las fachadas (hidrantes, torres de refrigeración, etc.) y el requisito de espesor de la capa base del nivel más bajo; fontanería a climatización: cuando la fuente de calor la suministre la especialidad de energía, formular la carga térmica total estimada y los parámetros de temperatura del agua secundaria (si el fluido caloportador es vapor, formular el modo de retorno de condensados); climatización a arquitectura: completar, mejorar y confirmar la superficie, las dimensiones y la altura libre de salas de refrigeración, salas de bombas, salas de calderas, estaciones de intercambio térmico, estaciones (salas) de regulación de gas, etc., las posiciones y superficies de los locales descentralizados de climatización y ventilación en cada nivel, las posiciones y dimensiones de los patinillos técnicos centrales, la posición y superficie preliminares de los conductos (superficies) de alivio de las salas de calderas, etc.; climatización a fontanería: los nombres de las salas de los puntos principales de consumo de agua y el consumo estimado, los nombres de las salas de los puntos principales de drenaje y la naturaleza del drenaje, el caudal total del agua de refrigeración, los requisitos de temperatura de entrada/salida y el número propuesto de grupos enfriadores; climatización a electricidad: los nombres de los equipos que requieren puesta a tierra de protección contra rayos y protección antiestática (salas de calderas de gasóleo/gas, chimeneas, tubos de venteo de gas, tuberías de gas combustible, etc.); electricidad a arquitectura: completar, mejorar y confirmar la superficie, las dimensiones y la altura libre de las subestaciones, salas de grupos electrógenos y otros locales técnicos, las posiciones y superficies de los locales de corrientes débiles y pequeños cuartos de distribución en cada nivel, las posiciones y dimensiones de los patinillos eléctricos principales, el requisito de espesor de los forjados, las posiciones y superficies de los conductos de admisión y extracción del enfriamiento de los grupos electrógenos; electricidad a fontanería: requisitos de protección contra incendios de los locales eléctricos especiales; sistemas inteligentes a arquitectura: completar, mejorar y confirmar la superficie, las dimensiones y la altura libre de los locales de acometida, locales de red, locales de centralitas telefónicas, locales de televisión por cable, locales de cobertura de señal móvil, salas de control de seguridad, locales de televisión por satélite, etc., las posiciones y dimensiones de los patinillos de corrientes débiles y de las salas de control de grandes salas de reuniones, y el requisito de espesor de los forjados",
        "pt": "Hidráulica para arquitetura: complementar, melhorar e confirmar os requisitos de área, dimensões e pé-direito de salas técnicas como salas de bombas, salas de tratamento de água, centrais de troca térmica, salas de água quente, salas de reservatórios, salas de válvulas de alarme, etc., as posições e dimensões dos shaft técnicos centrais, as dimensões e a localização em planta dos reservatórios, as posições e dimensões dos poços de captação do subsolo, as posições preliminares dos equipamentos que afetam a disposição interna e as fachadas (hidrantes, torres de resfriamento, etc.) e a exigência de espessura da camada de base do nível mais baixo; hidráulica para climatização: quando a fonte de calor for fornecida pela especialidade de energia, apresentar a carga térmica total estimada e os parâmetros de temperatura da água secundária (se o fluido térmico for vapor, apresentar o modo de retorno de condensado); climatização para arquitetura: complementar, melhorar e confirmar a área, as dimensões e o pé-direito de salas de refrigeração, salas de bombas, casas de caldeiras, centrais de troca térmica, estações (salas) de regulagem de gás, etc., as posições e áreas dos ambientes descentralizados de climatização e ventilação em cada pavimento, as posições e dimensões dos shaft técnicos centrais, a posição e a área preliminares dos dutos (superfícies) de alívio das casas de caldeiras, etc.; climatização para hidráulica: os nomes dos ambientes dos principais pontos de consumo de água e o consumo estimado, os nomes dos ambientes dos principais pontos de drenagem e a natureza da drenagem, a vazão total da água de resfriamento, os requisitos de temperatura de entrada/saída e o número proposto de chillers; climatização para elétrica: os nomes dos equipamentos que exigem aterramento contra raios e proteção antiestática (casas de caldeiras a óleo/gás, chaminés, tubos de vente de gás, tubulações de gás combustível, etc.); elétrica para arquitetura: complementar, melhorar e confirmar a área, as dimensões e o pé-direito de subestações, salas de geradores e outras salas técnicas, as posições e áreas dos ambientes de corrente fraca e pequenas salas de distribuição em cada pavimento, as posições e dimensões dos shaft elétricos principais, a exigência de espessura das lajes, as posições e áreas dos dutos de admissão e exaustão do resfriamento dos geradores; elétrica para hidráulica: requisitos de proteção contra incêndio dos ambientes elétricos especiais; sistemas inteligentes para arquitetura: complementar, melhorar e confirmar a área, as dimensões e o pé-direito de salas de entrada, salas de rede, salas de centrais telefônicas, salas de TV a cabo, salas de cobertura de sinal móvel, salas de controle de segurança, salas de TV via satélite, etc., as posições e dimensões dos shaft de corrente fraca e das salas de controle de grandes salas de reunião, e a exigência de espessura das lajes",
        "ru": "Водоснабжение — архитектуре: дополнить, уточнить и подтвердить требования к площадям, размерам и высотам помещений насосных, водоподготовки, теплообменных пунктов, горячего водоснабжения, водяных баков, пожарных клапанных и т.п., положение и размеры центральных шахт, размеры и плановую привязку баков, положение и размеры приямков в подвале, предварительное расположение оборудования, влияющего на внутреннюю планировку и фасады (пожарные гидранты, градирни и т.п.), требование к толщине выравнивающего слоя самого нижнего этажа; водоснабжение — отоплению и вентиляции: при подаче тепла от энергетической специальности представить расчётную суммарную тепловую нагрузку и параметры температуры вторичной воды (при паровом теплоносителе — способ возврата конденсата); отопление и вентиляция — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты холодильных, насосных, котельных, теплообменных пунктов, газорегуляторных пунктов (помещений) и т.п., расположение и площади децентрализованных вентиляционных камер на каждом этаже, положение и размеры центральных шахт, предварительное расположение и площадь предохранительных шахт (сбросных поверхностей) котельных и т.п.; отопление и вентиляция — водоснабжению: наименования помещений основных точек водоразбора и расчётные расходы воды, наименования помещений основных точек водоотведения и характер водоотведения, общий расход охлаждающей воды, требования к температурам на входе и выходе и планируемое количество холодильных машин; отопление и вентиляция — электроснабжению: наименования оборудования, требующего молниезащитного заземления и антистатической защиты (мазутно-газовые котельные, дымовые трубы, газовые свечи рассеивания, трубопроводы горючих газов и т.п.); электроснабжение — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты трансформаторных подстанций, дизельных и т.п., расположение и площади слаботочных помещений и распределительных ниш на каждом этаже, положение и размеры основных силовых шахт, требование к толщине плит перекрытий, расположение и площади шахт приточной и вытяжной вентиляции охлаждения генераторов; электроснабжение — водоснабжению: требования пожарной безопасности к специальным электротехническим помещениям; интеллектуальные системы — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений ввода, серверных, АТС, кабельного телевидения, покрытия сотовой связи, охранных диспетчерских, спутникового телевидения и т.п., положение и размеры слаботочных шахт, диспетчерских больших конференц-залов и требование к толщине плит перекрытий",
        "ar": "السباكة إلى العمارة: استكمال وتحسين وتأكيد متطلبات المساحة والأبعاد والارتفاع الصافي للغرف الفنية مثل غرف المضخات وغرف معالجة المياه ومحطات التبادل الحراري وغرف الماء الساخن وغرف خزانات المياه وغرف صمامات الإنذار وغيرها، ومواقع وأبعاد الأعمدة الفنية المركزية، وأبعاد وموقع خزانات المياه في المخطط، ومواقع وأبعاد أحواض التجميع في الطابق السفلي، والمواقع الأولية للمعدات التي تؤثر على التوزيع الداخلي والواجهات (صنابير الإطفاء وأبراج التبريد وغيرها)، ومتطلب سمك طبقة التأسيس لأدنى طابق؛ والسباكة إلى التدفئة والتهوية: عند توفير مصدر الحرارة من تخصص الطاقة، اقتراح الحمل الحراري الإجمالي التقديري ومعايير درجة حرارة الماء الثانوي (إذا كان وسيط التسخين بخاراً، اقتراح طريقة إرجاع المكثفات)؛ والتدفئة والتهوية إلى العمارة: استكمال وتحسين وتأكيد مساحة وأبعاد وارتفاع غرف التبريد وغرف المضخات وغرف الغلايات ومحطات التبادل الحراري ومحطات (غرف) تنظيم ضغط الغاز وغيرها، ومواقع ومساحات غرف التكييف والتهوية الموزعة في كل طابق، ومواقع وأبعاد الأعمدة الفنية المركزية، والموقع والمساحة الأوليان لأعمدة (أسطح) تفريغ الضغط في غرف الغلايات وغيرها؛ والتدفئة والتهوية إلى السباكة: أسماء غرف نقاط استهلاك الماء الرئيسية واستهلاك الماء التقديري، وأسماء غرف نقاط الصرف الرئيسية وطبيعة الصرف، ومعدل التدفق الإجمالي لمياه التبريد ومتطلبات درجة حرارة الدخول/الخروج وعدد وحدات التبريد المقترح؛ والتدفئة والتهوية إلى الكهرباء: أسماء المعدات التي تتطلب تأريض الحماية من الصواعق والحماية من الكهرباء الساكنة (غرف الغلايات النفطية/الغازية والمداخن وأنابيب تصريف الغاز وخطوط الغاز القابلة للاشتعال وغيرها)؛ والكهرباء إلى العمارة: استكمال وتحسين وتأكيد مساحة وأبعاد وارتفاع المحطات الفرعية وغرف المولدات وغيرها من الغرف الفنية، ومواقع ومساحات غرف التيارات الضعيفة وحجرات التوزيع الصغيرة في كل طابق، ومواقع وأبعاد أعمدة التيار القوي الرئيسية، ومتطلب سمك الألواح، ومواقع ومساحات أعمدة سحب وطرد هواء تبريد المولدات؛ والكهرباء إلى السباكة: متطلبات الحماية من الحرائق للغرف الكهربائية الخاصة؛ والأنظمة الذكية إلى العمارة: استكمال وتحسين وتأكيد مساحة وأبعاد وارتفاع غرف خطوط الدخول وغرف الشبكات وغرف مقاسم الهاتف وغرف التلفزيون الكبلي وغرف تغطية إشارات الهاتف المحمول وغرف التحكم الأمنية وغرف التلفزيون الفضائي وغيرها، ومواقع وأبعاد أعمدة التيارات الضعيفة وغرف التحكم في قاعات الاجتماعات الكبيرة، ومتطلب سمك الألواح",
        "it": "Impianti idrici ad architettura: integrare, migliorare e confermare i requisiti di area, dimensioni e altezza libera di locali tecnici quali sale pompe, sale trattamento acqua, stazioni di scambio termico, sale acqua calda, sale serbatoi, sale valvole di allarme, ecc., le posizioni e dimensioni dei vani tecnici centrali, le dimensioni e la posizione in pianta dei serbatoi, le posizioni e dimensioni dei pozzetti di raccolta in seminterrato, le posizioni preliminari delle apparecchiature che influiscono sulla disposizione interna e sulle facciate (idranti, torri di raffreddamento, ecc.) e il requisito di spessore dello strato di base del livello più basso; idrici a HVAC: quando la fonte di calore è fornita dalla disciplina energetica, proporre il carico termico totale stimato e i parametri di temperatura dell'acqua secondaria (se il fluido termovettore è vapore, proporre la modalità di ritorno della condensa); HVAC ad architettura: integrare, migliorare e confermare l'area, le dimensioni e l'altezza libera di sale refrigerazione, sale pompe, centrali termiche, stazioni di scambio termico, stazioni (sale) di regolazione del gas, ecc., le posizioni e aree dei locali decentralizzati di climatizzazione e ventilazione a ogni piano, le posizioni e dimensioni dei vani tecnici centrali, la posizione e l'area preliminari dei condotti (superfici) di scarico pressione delle centrali termiche, ecc.; HVAC a idrici: i nomi dei locali dei principali punti di prelievo acqua e il consumo stimato, i nomi dei locali dei principali punti di scarico e la natura dello scarico, la portata totale dell'acqua di raffreddamento, i requisiti di temperatura ingresso/uscita e il numero proposto di gruppi frigoriferi; HVAC a elettrici: i nomi delle apparecchiature che richiedono messa a terra contro i fulmini e protezione antistatica (centrali a olio/gas, camini, tubi di sfogo gas, tubazioni di gas combustibile, ecc.); elettrici ad architettura: integrare, migliorare e confermare l'area, le dimensioni e l'altezza libera di cabine di trasformazione, sale generatori e altri locali tecnici, le posizioni e aree dei locali debole corrente e dei piccoli locali di distribuzione a ogni piano, le posizioni e dimensioni dei vani elettrici principali, il requisito di spessore delle solette, le posizioni e aree dei condotti di aspirazione e scarico del raffreddamento dei generatori; elettrici a idrici: requisiti di protezione antincendio dei locali elettrici speciali; sistemi intelligenti ad architettura: integrare, migliorare e confermare l'area, le dimensioni e l'altezza libera di locali di ingresso linea, locali rete, locali centralini telefonici, locali TV via cavo, locali di copertura segnale mobile, sale controllo sicurezza, locali TV via satellite, ecc., le posizioni e dimensioni dei vani debole corrente e delle sale di controllo delle grandi sale riunioni, e il requisito di spessore delle solette",
        "vi": "Cấp thoát nước gửi kiến trúc: bổ sung hoàn thiện và xác nhận yêu cầu diện tích, kích thước, chiều cao thông thủy của phòng bơm, phòng xử lý nước, trạm trao đổi nhiệt, phòng nước nóng, phòng bể nước, phòng van báo động, v.v., vị trí kích thước giếng kỹ thuật tập trung, kích thước và định vị mặt bằng bể nước, vị trí kích thước hố thu nước tầng hầm, vị trí sơ bộ của thiết bị ảnh hưởng bố trí nội thất và mặt ngoài (họng cứu hỏa, tháp giải nhiệt, v.v.), yêu cầu chiều dày lớp lót tầng thấp nhất; cấp thoát nước gửi điều hòa: khi nguồn nhiệt do chuyên ngành năng lượng cung cấp thì nêu phụ tải nhiệt tổng ước tính và thông số nhiệt độ nước thứ cấp (khi môi chất là hơi nước thì nêu phương thức hồi nước ngưng); điều hòa gửi kiến trúc: bổ sung hoàn thiện xác nhận diện tích, kích thước, chiều cao thông thủy phòng lạnh, phòng bơm, phòng lò hơi, trạm trao đổi nhiệt, trạm (phòng) điều áp gas, v.v., vị trí diện tích phòng điều hòa thông gió phân tán từng tầng, vị trí kích thước giếng kỹ thuật tập trung, vị trí diện tích sơ bộ giếng (bề mặt) thoát áp của phòng lò hơi, v.v.; điều hòa gửi cấp thoát nước: tên phòng điểm dùng nước chính và lượng nước dùng ước tính, tên phòng điểm thoát nước chính và tính chất thoát nước, tổng lưu lượng nước làm mát, yêu cầu nhiệt độ nước vào/ra và số máy lạnh dự kiến; điều hòa gửi điện: tên thiết bị cần tiếp đất chống sét và chống tĩnh điện (phòng lò hơi dầu khí, ống khói, ống xả khí gas, đường ống khí cháy được, v.v.); điện gửi kiến trúc: bổ sung hoàn thiện xác nhận diện tích, kích thước, chiều cao thông thủy trạm biến áp, phòng máy phát, v.v., vị trí diện tích phòng điện yếu và phòng phân phối nhỏ từng tầng, vị trí kích thước giếng điện mạnh chính, yêu cầu chiều dày sàn, vị trí diện tích giếng hút/ thải gió làm mát máy phát; điện gửi cấp thoát nước: yêu cầu phòng cháy của phòng điện đặc biệt; hệ thống thông minh gửi kiến trúc: bổ sung hoàn thiện xác nhận diện tích, kích thước, chiều cao thông thủy phòng cáp vào, phòng mạng, phòng tổng đài điện thoại, phòng truyền hình cáp, phòng phủ sóng di động, phòng điều khiển an ninh, phòng truyền hình vệ tinh, v.v., vị trí kích thước giếng điện yếu, phòng điều khiển phòng họp lớn, yêu cầu chiều dày sàn",
        "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น",
    },
    "水、暖、电提结构：位于楼板上的机房设备和冷却塔的结构荷载；水、暖提电：各强电设备的位置及其电气安装容量（消防设备和平时使用的设备分别提出）": {
        "en": "Plumbing, HVAC and electrical to structural: structural loads of equipment rooms on floor slabs and of cooling towers; plumbing and HVAC to electrical: positions of all high-voltage/strong-current equipment and their electrical installation capacities (fire-protection equipment and normally used equipment shall be stated separately)",
        "ja": "水・暖・電→構造：スラブ上の機械室設備と冷却塔の構造荷重；水・暖→電：各強電設備の位置およびその電気設備容量（消防設備と常用設備は分けて提出する）",
        "ko": "급배수·난방·전기→구조: 슬래브 위 기계실 설비와 냉각탑의 구조 하중; 급배수·난방→전기: 각 강전 설비의 위치 및 전기 설비 용량(소방 설비와 평상시 사용 설비는 각각 제출)",
        "fr": "Plomberie, CVC et électricité vers structure : charges structurelles des équipements de locaux techniques sur dalles et des tours de refroidissement ; plomberie et CVC vers électricité : emplacements des équipements de courant fort et leurs capacités électriques installées (les équipements de protection incendie et les équipements d'usage courant sont indiqués séparément)",
        "de": "Sanitär, Heizung-Lüftung und Elektro an Tragwerk: Strukturlasten der Geräte in Technikräumen auf Deckenplatten und von Kühltürmen; Sanitär und Heizung-Lüftung an Elektro: Positionen aller Starkstromgeräte und ihre elektrischen Anschlussleistungen (Brandschutzgeräte und normal genutzte Geräte sind getrennt anzugeben)",
        "es": "Fontanería, climatización y electricidad a estructuras: cargas estructurales de los equipos de locales técnicos sobre losas y de las torres de refrigeración; fontanería y climatización a electricidad: posiciones de todos los equipos de corriente fuerte y sus capacidades eléctricas instaladas (los equipos de protección contra incendios y los de uso habitual se indican por separado)",
        "pt": "Hidráulica, climatização e elétrica para estruturas: cargas estruturais dos equipamentos de salas técnicas sobre lajes e das torres de resfriamento; hidráulica e climatização para elétrica: posições de todos os equipamentos de corrente forte e suas capacidades elétricas instaladas (equipamentos de proteção contra incêndio e equipamentos de uso normal devem ser indicados separadamente)",
        "ru": "Водоснабжение, отопление и вентиляция, электроснабжение — конструкциям: конструктивные нагрузки от оборудования технических помещений на плитах перекрытий и от градирен; водоснабжение и отопление и вентиляция — электроснабжению: расположение всего силового электрооборудования и его установленная электрическая мощность (противопожарное оборудование и оборудование обычного использования указываются раздельно)",
        "ar": "السباكة والتدفئة والتهوية والكهرباء إلى الإنشاءات: الأحمال الإنشائية لمعدات الغرف الفنية على الألواح ولأبراج التبريد؛ والسباكة والتدفئة والتهوية إلى الكهرباء: مواقع جميع معدات التيار القوي وسعاتها الكهربائية المركبة (تُذكر معدات الحماية من الحرائق ومعدات الاستخدام العادي بشكل منفصل)",
        "it": "Impianti idrici, HVAC ed elettrici a strutture: carichi strutturali delle apparecchiature dei locali tecnici su solette e delle torri di raffreddamento; impianti idrici e HVAC a elettrici: posizioni di tutte le apparecchiature di forza e relative potenze elettriche installate (le apparecchiature antincendio e quelle di uso normale vanno indicate separatamente)",
        "vi": "Cấp thoát nước, điều hòa, điện gửi kết cấu: tải trọng kết cấu của thiết bị phòng máy trên sàn và tháp giải nhiệt; cấp thoát nước, điều hòa gửi điện: vị trí các thiết bị điện mạnh và công suất điện lắp đặt của chúng (thiết bị chữa cháy và thiết bị sử dụng thường xuyên nêu riêng)",
        "th": "ประปา ระบบปรับอากาศ และไฟฟ้า ส่งให้โครงสร้าง: น้ำหนักบรรทุกโครงสร้างของอุปกรณ์ห้องเครื่องบนพื้นแผ่นและหอหล่อเย็น ประปาและระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งของอุปกรณ์ไฟฟ้าแรงทั้งหมดและกำลังไฟฟ้าติดตั้ง (อุปกรณ์ดับเพลิงและอุปกรณ์ใช้งานปกติต้องแจ้งแยกกัน)",
    },
    "水、暖、电、智能化提结构：最大长边尺寸或直径≥1000mm的承重结构墙和楼板的预留洞初步位置；水、暖、电提总图：各市政管线（给水、排水、热力、燃气、电力、电信、有线等）在本楼的接口位置要求，主要室外管线布置及主要相关构筑物（化粪池、隔油池、水表井、水泵结合器、阀门井、管线检修井、烟囱、室外箱式变压器等）初步位置，有燃油系统时提出室外储油罐占用面积、初步位置和相关技术要求": {
        "en": "Plumbing, HVAC, electrical and intelligent systems to structural: the preliminary positions of reserved openings in load-bearing structural walls and floor slabs with a maximum long-side dimension or diameter of ≥1000mm; plumbing, HVAC and electrical to site planning: the interface position requirements of various municipal pipelines (water supply, drainage, heating, gas, power, telecom, cable TV, etc.) entering this building, the preliminary positions of main outdoor pipeline layouts and main related structures (septic tanks, grease interceptors, water meter wells, fire pump connections, valve wells, pipeline inspection wells, chimneys, outdoor pad-mounted transformers, etc.); when there is an oil fuel system, raise the occupied area, preliminary position and related technical requirements of the outdoor fuel storage tank",
        "ja": "水・暖・電・智能化から構造専門へ：最大長辺寸法または直径が1000mm以上の耐力構造壁と床スラブの予備開口の初期位置；水・暖・電から総図専門へ：各市政配管（給水・排水・熱力・ガス・電力・通信・有線等）の本棟への接続位置要求、主要屋外配管の配置および主要関連構築物（化糞池・グリーストラップ・水道メーター桝・ポンプ接合器・バルブ桝・管路点検桝・煙突・屋外キュービクル変圧器等）の初期位置；燃油システムがある場合、屋外燃料貯蔵タンクの占有面積・初期位置・関連技術要求を提出する",
        "ko": "급배수·난방·전기·지능화가 구조에: 최대 장변 치수 또는 직경이 1000mm 이상인 내력 구조벽과 슬래브의 예비 구멍 초기 위치; 급배수·난방·전기가 총도에: 각 시정 배관(급수·배수·열력·가스·전력·통신·유선 등)의 본 건물 접속 위치 요구, 주요 옥외 배관 배치 및 주요 관련 구조물(정화조·기름분리조·수도계량기 우물·소방펌프접합기·밸브 우물·관로 점검 우물·굴뚝·옥외 박스변압기 등)의 초기 위치; 유류 시스템이 있을 때 옥외 유류 저장탱크 점유 면적·초기 위치·관련 기술 요구를 제출합니다",
        "fr": "Plomberie, CVC, électricité et systèmes intelligents à structure : les positions préliminaires des réservations dans les murs porteurs et les dalles dont la plus grande dimension ou le diamètre est ≥ 1000 mm ; plomberie, CVC et électricité à aménagement du site : les exigences de position des raccordements des divers réseaux municipaux (eau potable, assainissement, chauffage urbain, gaz, électricité, télécommunications, câble, etc.) au bâtiment, les positions préliminaires des principaux tracés de réseaux extérieurs et des principales ouvrages associés (fosses septiques, séparateurs de graisses, regards de compteurs d'eau, raccords de pompes incendie, regards de vannes, regards de visite, cheminées, postes de transformation extérieurs, etc.) ; en cas de système au fioul, formuler la surface occupée, la position préliminaire et les exigences techniques du réservoir de stockage de fioul extérieur",
        "de": "Sanitär, Heizung-Lüftung, Elektro und Gebäudeautomation an Tragwerk: die vorläufigen Positionen von Durchbrüchen in tragenden Wänden und Geschossdecken mit maximaler Längsseite oder Durchmesser ≥ 1000 mm; Sanitär, Heizung-Lüftung und Elektro an Standortplanung: die Anforderungen an die Anschlusspositionen der verschiedenen städtischen Leitungen (Wasserversorgung, Entwässerung, Fernwärme, Gas, Strom, Telekom, Kabel usw.) am Gebäude, die vorläufigen Positionen der wichtigsten Außenleitungsführungen und der wichtigsten zugehörigen Bauwerke (Klärgruben, Fettabscheider, Wasserzähler-Schächte, Feuerlöschpumpenanschlüsse, Schieber-Schächte, Leitungsrevisionsschächte, Schornsteine, Ortsnetzstationen im Freien usw.); bei einem Ölfeuerungssystem: belegte Fläche, vorläufige Position und technische Anforderungen des Außen-Öllagerbehälters angeben",
        "es": "Fontanería, climatización, electricidad y sistemas inteligentes a estructuras: las posiciones preliminares de los huecos reservados en muros portantes y forjados cuya mayor dimensión o diámetro sea ≥ 1000 mm; fontanería, climatización y electricidad a urbanización: los requisitos de posición de las conexiones de los distintos redes municipales (abastecimiento, saneamiento, calefacción, gas, electricidad, telecomunicaciones, cable, etc.) al edificio, las posiciones preliminares de los principales trazados de redes exteriores y de las principales obras asociadas (fosas sépticas, separadores de grasas, arquetas de contadores de agua, conexiones de bombas de incendio, arquetas de válvulas, arquetas de registro, chimeneas, centros de transformación exteriores, etc.); en caso de sistema de combustible líquido, formular la superficie ocupada, la posición preliminar y los requisitos técnicos del depósito exterior de almacenamiento de combustible",
        "pt": "Hidráulica, climatização, elétrica e sistemas inteligentes para estruturas: as posições preliminares das aberturas reservadas em paredes estruturais e lajes cuja maior dimensão ou diâmetro seja ≥ 1000 mm; hidráulica, climatização e elétrica para urbanismo: os requisitos de posição das conexões das diversas redes municipais (abastecimento, drenagem, aquecimento, gás, eletricidade, telecomunicações, cabo, etc.) ao edifício, as posições preliminares dos principais traçados de redes externas e das principais obras associadas (fossas sépticas, separadores de gordura, poços de hidrômetros, conexões de bombas de incêndio, poços de válvulas, poços de inspeção, chaminés, subestações externas, etc.); no caso de sistema de óleo combustível, apresentar a área ocupada, a posição preliminar e os requisitos técnicos do tanque externo de armazenamento de combustível",
        "ru": "Водоснабжение, отопление и вентиляция, электроснабжение и интеллектуальные системы — конструктивной специальности: предварительное расположение отверстий в несущих стенах и плитах перекрытий с наибольшим размером стороны или диаметром ≥ 1000 мм; водоснабжение, отопление и вентиляция, электроснабжение — специальности генплана: требования к местам присоединения городских сетей (водопровод, канализация, теплоснабжение, газ, электроэнергия, связь, кабельное ТВ и т.п.) к зданию, предварительное расположение основных наружных сетей и основных связанных сооружений (септики, жироуловители, водомерные колодцы, соединительные головки пожарных насосов, задвижечные колодцы, смотровые колодцы, дымовые трубы, наружные блочные трансформаторные подстанции и т.п.); при наличии системы на жидком топливе представить занимаемую площадь, предварительное расположение и технические требования к наружному резервуару хранения топлива",
        "ar": "السباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية إلى الإنشاءات: المواقع الأولية للفتحات المخصصة في الجدران الإنشائية الحاملة والألواح التي يبلغ أكبر بعد أو قطرها ≥ 1000 مم؛ والسباكة والتدفئة والتهوية والكهرباء إلى تخطيط الموقع: متطلبات مواقع توصيل الشبكات البلدية المختلفة (إمداد المياه والصرف والتدفئة والغاز والكهرباء والاتصالات والكابل وغيرها) بالمبنى، والمواقع الأولية لتخطيطات الشبكات الخارجية الرئيسية والمنشآت المرتبطة الرئيسية (خزانات الصرف الصحي وفواصل الدهون وآبار عدادات المياه ووصلات مضخات الإطفاء وآبار الصمامات وآبار فحص الخطوط والمداخن ومحولات الصناديق الخارجية وغيرها)؛ وعند وجود نظام وقود سائل، اقتراح المساحة المشغولة والموقع الأولي والمتطلبات الفنية لخزان تخزين الوقود الخارجي",
        "it": "Idrici, HVAC, elettrici e sistemi intelligenti a strutture: le posizioni preliminari delle aperture riservate nelle pareti strutturali portanti e nelle solette con dimensione massima del lato o diametro ≥ 1000 mm; idrici, HVAC ed elettrici a pianificazione del sito: i requisiti di posizione dei collegamenti delle varie reti comunali (acquedotto, fognatura, teleriscaldamento, gas, energia elettrica, telecomunicazioni, cavo, ecc.) all'edificio, le posizioni preliminari dei principali tracciati di reti esterne e delle principali opere correlate (fosse settiche, separatori di grassi, pozzetti contatori acqua, attacchi pompe antincendio, pozzetti valvole, pozzetti di ispezione, camini, cabine di trasformazione esterne, ecc.); in caso di sistema a olio combustibile, proporre l'area occupata, la posizione preliminare e i requisiti tecnici del serbatoio esterno di stoccaggio del combustibile",
        "vi": "Cấp thoát nước, điều hòa, điện, hệ thống thông minh gửi kết cấu: vị trí sơ bộ các lỗ chờ trên tường chịu lực và sàn có cạnh dài lớn nhất hoặc đường kính ≥1000mm; cấp thoát nước, điều hòa, điện gửi tổng mặt bằng: yêu cầu vị trí đấu nối các tuyến hạ tầng kỹ thuật (cấp nước, thoát nước, nhiệt, gas, điện, viễn thông, truyền hình cáp, v.v.) vào tòa nhà, vị trí sơ bộ bố trí đường ống ngoài nhà chính và các công trình liên quan chính (bể tự hoại, bể tách dầu mỡ, giếng đồng hồ nước, đầu nối bơm cứu hỏa, giếng van, giếng kiểm tra đường ống, ống khói, trạm biến áp ngoài trời, v.v.); khi có hệ thống nhiên liệu dầu thì nêu diện tích chiếm chỗ, vị trí sơ bộ và yêu cầu kỹ thuật liên quan của bể chứa dầu ngoài trời",
        "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งเบื้องต้นของช่องเผื่อในผนังโครงสร้างรับน้ำหนักและพื้นที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ประปา ระบบปรับอากาศ และไฟฟ้าส่งให้การวางผังบริเวณ: ข้อกำหนดตำแหน่งเชื่อมต่อท่อสาธารณูปโภคต่างๆ (น้ำประปา ระบายน้ำ ความร้อน แก๊ส ไฟฟ้า โทรคมนาคม เคเบิล ฯลฯ) เข้าอาคาร ตำแหน่งเบื้องต้นของการวางท่อภายนอกหลักและสิ่งก่อสร้างที่เกี่ยวข้องหลัก (บ่อเกรอะ ถังแยกไขมัน บ่อมิเตอร์น้ำ จุดต่อปั๊มดับเพลิง บ่อวาล์ว บ่อตรวจสอบท่อ ปล่องไฟ หม้อแปลงชนิดตู้นอกอาคาร ฯลฯ) เมื่อมีระบบน้ำมันเชื้อเพลิง ให้เสนอพื้นที่ใช้สอย ตำแหน่งเบื้องต้น และข้อกำหนดทางเทคนิคที่เกี่ยวข้องของถังเก็บน้ำมันภายนอก",
    },
    "结构专业提出：楼面、屋面结构梁板平面图；与建筑有联系的挡土墙结构形式（提建筑、总图专业）；设计配合所必需的特殊结构构造做法或详图（提建筑、总图专业）。接受专业：建筑、水、暖、电、经济": {
        "en": "Proposed by the structural discipline: structural beam and slab plans of floors and roofs; structural forms of retaining walls related to the architecture (to the architectural and site planning disciplines); special structural construction practices or details necessary for design coordination (to the architectural and site planning disciplines). Receiving disciplines: architectural, plumbing, HVAC, electrical, economics",
        "ja": "構造専門が提出する：床・屋根の構造梁伏図・スラブ図；建築と関連する擁壁の構造形式（建築・総図専門に提出）；設計配合に必要な特殊構造の仕様または詳細図（建築・総図専門に提出）。受取専門：建築、水、暖、電、経済",
        "ko": "구조 전문이 제출합니다: 바닥·지붕 구조 보·슬래브 평면도; 건축과 관련된 옹벽 구조 형식(건축·총도 전문에 제출); 설계 배합에 필요한 특수 구조 공법 또는 상세도(건축·총도 전문에 제출). 수신 전문: 건축, 급배수, 난방, 전기, 경제",
        "fr": "Présenté par la discipline structure : plans de structure (poutres et dalles) des planchers et des toitures ; formes structurelles des murs de soutènement liés à l'architecture (aux disciplines architecture et aménagement du site) ; pratiques constructives ou détails de structures spéciales nécessaires à la coordination (aux disciplines architecture et aménagement du site). Disciplines destinataires : architecture, plomberie, CVC, électricité, économie",
        "de": "Vorgelegt von der Tragwerksdisziplin: Strukturpläne (Träger und Platten) der Geschossdecken und Dächer; Tragwerksformen der mit der Architektur zusammenhängenden Stützmauern (an die Architektur- und Standortplanungsdisziplin); besondere konstruktive Ausführungsarten oder Detailzeichnungen, die für die Fachkoordination erforderlich sind (an die Architektur- und Standortplanungsdisziplin). Empfangende Disziplinen: Architektur, Sanitär, Heizung-Lüftung, Elektro, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de estructuras: planos estructurales de vigas y losas de forjados y cubiertas; formas estructurales de los muros de contención relacionados con la arquitectura (a las especialidades de arquitectura y urbanización); prácticas constructivas o detalles estructurales especiales necesarios para la coordinación (a las especialidades de arquitectura y urbanización). Especialidades receptoras: arquitectura, fontanería, climatización, electricidad, economía",
        "pt": "Apresentado pela especialidade de estruturas: plantas estruturais de vigas e lajes de pisos e coberturas; formas estruturais dos muros de arrimo relacionados à arquitetura (às especialidades de arquitetura e urbanismo); práticas construtivas ou detalhes estruturais especiais necessários à coordenação (às especialidades de arquitetura e urbanismo). Especialidades receptoras: arquitetura, hidráulica, climatização, elétrica, economia",
        "ru": "Представляется конструктивной специальностью: планы несущих балок и плит перекрытий и покрытий; конструктивные формы подпорных стен, связанных с архитектурой (для архитектурной специальности и генплана); особые конструктивные решения или детали, необходимые для координации проектирования (для архитектурной специальности и генплана). Принимающие специальности: архитектура, водоснабжение, отопление и вентиляция, электроснабжение, экономика",
        "ar": "يقدمه تخصص الإنشاءات: مخططات الكمرات والألواح الإنشائية للأرضيات والأسطح؛ والأشكال الإنشائية للجدران الاستنادية المرتبطة بالعمارة (لتخصصي العمارة وتخطيط الموقع)؛ وأساليب الإنشاء الخاصة أو التفاصيل اللازمة للتنسيق (لتخصصي العمارة وتخطيط الموقع). التخصصات المستلمة: العمارة، السباكة، التدفئة والتهوية، الكهرباء، الاقتصاد",
        "it": "Presentato dalla disciplina strutturale: piante strutturali di travi e solette di solai e coperture; forme strutturali dei muri di contenimento collegati all'architettura (alle discipline architettonica e di pianificazione del sito); pratiche costruttive o dettagli strutturali speciali necessari al coordinamento (alle discipline architettonica e di pianificazione del sito). Discipline destinatarie: architettura, impianti idrici, HVAC, elettrici, economia",
        "vi": "Do chuyên ngành kết cấu đưa ra: mặt bằng kết cấu dầm, sàn của sàn và mái; hình thức kết cấu tường chắn liên quan đến kiến trúc (gửi chuyên ngành kiến trúc, tổng mặt bằng); biện pháp cấu tạo hoặc chi tiết kết cấu đặc biệt cần thiết cho phối hợp thiết kế (gửi chuyên ngành kiến trúc, tổng mặt bằng). Các chuyên ngành nhận: kiến trúc, cấp thoát nước, điều hòa không khí, điện, kinh tế",
        "th": "เสนอโดยสาขาโครงสร้าง: แผนผังโครงสร้างคานและพื้นของพื้นอาคารและหลังคา; รูปแบบโครงสร้างของกำแพงกันดินที่เกี่ยวข้องกับสถาปัตยกรรม (ส่งให้สาขาสถาปัตยกรรมและการวางผังบริเวณ); วิธีก่อสร้างโครงสร้างพิเศษหรือแบบขยายที่จำเป็นสำหรับการประสานงานออกแบบ (ส่งให้สาขาสถาปัตยกรรมและการวางผังบริเวณ) สาขาที่รับ: สถาปัตยกรรม ประปา ระบบปรับอากาศ ไฟฟ้า เศรษฐศาสตร์",
    },
    "项目经理组织、设计主持人主持、各专业负责人参加（第三阶段开始日）：第二阶段情况总结；检查《专业配合进度表》执行情况，针对各专业配合资料进行协调，解决上阶段配合过程中的问题；安排下一阶段工作": {
        "en": "Organized by the project manager, chaired by the design director, attended by the heads of all disciplines (on the start date of Stage 3): summarize the work of Stage 2; check the implementation of the Professional Coordination Schedule; coordinate the coordination materials of all disciplines and resolve the problems encountered in the previous stage; arrange the work of the next stage",
        "ja": "プロジェクトマネージャーが組織し、設計主催者が司会し、各専門担当責任者が参加する（第3段階開始日）：第2段階の状況を総括する；『専門配合進度表』の実行状況を点検し、各専門の配合資料について調整を行い、前段階の配合過程の問題を解決する；次段階の作業を手配する",
        "ko": "프로젝트 매니저가 조직하고 설계 책임자가 주재하며 각 전문 담당자가 참석합니다(3단계 시작일): 2단계 상황을 총결합니다; 《전문배합진도표》 실행 상황을 점검하고 각 전문의 배합 자료를 조정하며 이전 단계 배합 과정의 문제를 해결합니다; 다음 단계 작업을 배치합니다",
        "fr": "Organisé par le chef de projet, présidé par le directeur de conception, avec la participation des responsables de toutes les disciplines (au début de la phase 3) : bilan de la phase 2 ; contrôle de l'exécution du Calendrier de coordination professionnelle, coordination des documents de toutes les disciplines et résolution des problèmes rencontrés au cours de la phase précédente ; organisation des travaux de la phase suivante",
        "de": "Organisiert durch den Projektleiter, geleitet durch den Entwurfsleiter, mit Teilnahme der Verantwortlichen aller Fachdisziplinen (am Beginn von Stufe 3): Zusammenfassung der Stufe 2; Überprüfung der Umsetzung des Fachkoordinationszeitplans, Koordination der Fachunterlagen und Lösung der in der vorherigen Stufe aufgetretenen Probleme; Einteilung der Arbeiten der nächsten Stufe",
        "es": "Organizado por el jefe de proyecto, presidido por el director de diseño, con participación de los responsables de todas las especialidades (al inicio de la fase 3): resumen de la fase 2; comprobación de la ejecución del Cronograma de coordinación profesional, coordinación de los documentos de todas las especialidades y resolución de los problemas surgidos en la fase anterior; organización de los trabajos de la fase siguiente",
        "pt": "Organizado pelo gerente do projeto, presidido pelo diretor de projeto, com a participação dos responsáveis de todas as especialidades (no início da fase 3): resumo da fase 2; verificação da execução do Cronograma de Coordenação Profissional, coordenação dos documentos de todas as especialidades e resolução dos problemas ocorridos na fase anterior; organização dos trabalhos da fase seguinte",
        "ru": "Организует руководитель проекта, ведёт главный архитектор проекта, участвуют ответственные по всем специальностям (на дату начала 3-го этапа): подведение итогов 2-го этапа; проверка выполнения Графика профессиональной координации, координация материалов всех специальностей и решение проблем, возникших в ходе предыдущего этапа; организация работ следующего этапа",
        "ar": "ينظمه مدير المشروع، ويرأسه مدير التصميم، ويشارك فيه رؤساء جميع التخصصات (في يوم بدء المرحلة الثالثة): تلخيص نتائج المرحلة الثانية؛ والتحقق من تنفيذ جدول التنسيق المهني، والتنسيق بشأن مواد التنسيق لجميع التخصصات وحل المشكلات التي نشأت خلال المرحلة السابقة؛ وترتيب أعمال المرحلة التالية",
        "it": "Organizzato dal project manager, presieduto dal direttore della progettazione, con la partecipazione dei responsabili di tutte le discipline (all'inizio della fase 3): sintesi della fase 2; verifica dell'attuazione del Cronoprogramma di coordinamento professionale, coordinamento dei documenti di tutte le discipline e risoluzione dei problemi emersi nella fase precedente; organizzazione dei lavori della fase successiva",
        "vi": "Do quản lý dự án tổ chức, giám đốc thiết kế chủ trì, các trưởng bộ môn tham dự (ngày bắt đầu giai đoạn 3): tổng kết tình hình giai đoạn 2; kiểm tra việc thực hiện Bảng tiến độ phối hợp chuyên môn, phối hợp tài liệu của các chuyên ngành và giải quyết các vấn đề trong quá trình phối hợp giai đoạn trước; bố trí công việc giai đoạn tiếp theo",
        "th": "จัดโดยผู้จัดการโครงการ โดยผู้อำนวยการฝ่ายออกแบบเป็นประธาน และหัวหน้าทุกสาขาวิชาเข้าร่วม (วันเริ่มต้นระยะที่ 3): สรุปสถานการณ์ระยะที่ 2; ตรวจสอบการดำเนินการตามตารางการประสานงานวิชาชีพ ประสานงานเอกสารของทุกสาขาและแก้ไขปัญหาที่เกิดขึ้นในระหว่างการประสานงานระยะก่อนหน้า; จัดงานในระยะถัดไป",
    },
    "建筑专业提出完整的电子版文件，作为各专业初步设计正式绘图用底图。内容（在第一版建筑作业图基础上）：全面完整反映第二阶段各专业提资料、协调配合成果；完善防火分区图、人防平面及口部布置与做法、全部房间名称、建筑专业本身尺寸标注；对特定房间提供必要详图放大图（客房、住宅单元放大图、屋面及外墙构造详图）；外门、外窗尺寸及性能要求（光学、热工性能）；对预留竖井标注使用专业。接受专业：总图、结构、水、暖、电、智能化、经济": {
        "en": "The architectural discipline submits complete electronic files as the base drawings for the formal drafting of the preliminary design by all disciplines. Contents (on the basis of the first version of the architectural working drawings): comprehensively and completely reflect the materials submitted by all disciplines and the coordination results of Stage 2; improve the fire compartmentation plans, civil air defense plans and the layout and practices of openings, all room names, and the architectural discipline's own dimensioning; provide necessary enlarged detail drawings for specific rooms (enlarged drawings of guest rooms and residential units, construction detail drawings of roofs and external walls); the dimensions and performance requirements of external doors and windows (optical and thermal performance); annotate the using disciplines on reserved riser shafts. Receiving disciplines: site planning, structural, plumbing, HVAC, electrical, intelligent systems, economics",
        "ja": "建築専門が完全な電子版ファイルを提出し、各専門の基本設計正式製図用の下図とする。内容（第一版建築作業図を基礎とする）：第二段階の各専門提出資料・調整配合成果を全面的かつ完全に反映する；防火区画図、人防平面及び口部配置・仕様、全室名、建築専門自身の寸法記入を完成させる；特定の部屋に必要な詳細拡大図（客室・住宅ユニット拡大図、屋根・外壁構造詳細図）を提供する；外扉・外窓の寸法と性能要求（光学・熱工性能）；予備シャフトに使用専門を注記する。受取専門：総図、構造、水、暖、電、智能化、経済",
        "ko": "건축 전문이 완전한 전자파일을 제출하여 각 전문의 기본설계 정식 도면 작성을 위한 밑그림으로 사용합니다. 내용(제1판 건축 작업도를 기반으로): 제2단계 각 전문 제출 자료와 조정·배합 성과를 전면적이고 완전하게 반영합니다; 방화구획도, 인방 평면 및 출입구 배치·시공법, 전체 실명, 건축 전문 자체의 치수 기입을 완성합니다; 특정 실에는 필요한 상세 확대도를 제공합니다(객실·주거 유닛 확대도, 지붕·외벽 구조 상세도); 외문·외창의 치수와 성능 요구(광학·열공 성능); 예비 샤프트에 사용 전문을 표기합니다. 수신 전문: 총도, 구조, 급배수, 난방, 전기, 지능화, 경제",
        "fr": "La discipline architecture soumet des fichiers électroniques complets comme fonds de plan pour le dessin formel du projet d'esquisse de toutes les disciplines. Contenu (sur la base de la première version des plans de travail d'architecture) : refléter de manière globale et complète les documents soumis par toutes les disciplines et les résultats de coordination de la phase 2 ; compléter les plans de compartimentage coupe-feu, les plans de défense civile et la disposition et les pratiques des ouvertures, tous les noms des pièces et le dimensionnement propre à la discipline architecture ; fournir les agrandissements de détail nécessaires pour les pièces spécifiques (agrandissements des chambres et des unités résidentielles, détails constructifs des toitures et des murs extérieurs) ; les dimensions et exigences de performance des portes et fenêtres extérieures (performances optiques et thermiques) ; annoter les disciplines utilisatrices sur les gaines réservées. Disciplines destinataires : aménagement du site, structure, plomberie, CVC, électricité, systèmes intelligents, économie",
        "de": "Die Architekturdisziplin liefert vollständige elektronische Dateien als Vorlagen für die formelle Zeichnung der Vorplanung aller Fachdisziplinen. Inhalt (auf der Grundlage der ersten Fassung der Architektur-Arbeitszeichnungen): die in Stufe 2 eingereichten Unterlagen aller Disziplinen und die Koordinationsergebnisse umfassend und vollständig widerspiegeln; die Brandabschnittspläne, die Schutzraumpäne sowie Anordnung und Ausführung der Öffnungen, alle Raumnamen und die eigene Bemaßung der Architekturdisziplin vervollständigen; für bestimmte Räume notwendige Detailvergrößerungen liefern (Vergrößerungen von Gästezimmern und Wohnungseinheiten, Konstruktionsdetails von Dächern und Außenwänden); Abmessungen und Leistungsanforderungen der Außentüren und -fenster (optische und thermische Leistung); an reservierten Schächten die nutzenden Disziplinen vermerken. Empfangende Disziplinen: Standortplanung, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation, Wirtschaftlichkeit",
        "es": "La especialidad de arquitectura presenta archivos electrónicos completos como base para el dibujo formal del diseño preliminar de todas las especialidades. Contenido (sobre la base de la primera versión de los planos de trabajo de arquitectura): reflejar de forma global y completa los documentos presentados por todas las especialidades y los resultados de coordinación de la fase 2; completar los planos de compartimentación contra incendios, los planos de defensa civil y la disposición y prácticas de las aberturas, todos los nombres de los locales y el acotado propio de la especialidad de arquitectura; proporcionar los alzados ampliados de detalle necesarios para locales específicos (ampliados de habitaciones y unidades residenciales, detalles constructivos de cubiertas y muros exteriores); las dimensiones y requisitos de prestaciones de puertas y ventanas exteriores (prestaciones ópticas y térmicas); anotar las especialidades usuarias en los patinillos reservados. Especialidades receptoras: urbanización, estructuras, fontanería, climatización, electricidad, sistemas inteligentes, economía",
        "pt": "A especialidade de arquitetura apresenta arquivos eletrônicos completos como base para o desenho formal do projeto preliminar de todas as especialidades. Conteúdo (com base na primeira versão dos desenhos de trabalho de arquitetura): refletir de forma abrangente e completa os documentos apresentados por todas as especialidades e os resultados de coordenação da fase 2; completar os desenhos de compartimentação contra incêndio, os desenhos de defesa civil e a disposição e práticas das aberturas, todos os nomes dos ambientes e a cotagem própria da especialidade de arquitetura; fornecer ampliações de detalhe necessárias para ambientes específicos (ampliações de quartos de hóspedes e unidades residenciais, detalhes construtivos de coberturas e paredes externas); as dimensões e requisitos de desempenho de portas e janelas externas (desempenho óptico e térmico); anotar as especialidades usuárias nos shaft reservados. Especialidades receptoras: urbanismo, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes, economia",
        "ru": "Архитектурная специальность представляет полные электронные файлы в качестве подосновы для официального оформления чертежей стадии «проект» всеми специальностями. Содержание (на основе первой версии рабочих архитектурных чертежей): всесторонне и полно отразить материалы, представленные всеми специальностями на 2-м этапе, и результаты координации; завершить схемы противопожарных отсеков, планы гражданской обороны с расположением и устройством входов, все наименования помещений и простановку размеров самой архитектурной специальности; для отдельных помещений представить необходимые увеличенные детальные чертежи (укрупнённые планы гостиничных номеров и квартир, конструктивные детали кровли и наружных стен); размеры и требования к характеристикам наружных дверей и окон (оптические и теплотехнические характеристики); на зарезервированных шахтах указать использующие специальности. Принимающие специальности: генплан, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы, экономика",
        "ar": "يقدم تخصص العمارة ملفات إلكترونية كاملة كأساس للرسم الرسمي للتصميم الأولي لجميع التخصصات. المحتوى (استناداً إلى النسخة الأولى من مخططات العمل المعمارية): عكس المواد المقدمة من جميع التخصصات ونتائج التنسيق في المرحلة الثانية بشكل شامل وكامل؛ واستكمال مخططات تقسيمات الحماية من الحرائق، ومخططات الدفاع المدني وترتيب وأساليب الفتحات، وجميع أسماء الغرف، وتحديد الأبعاد الخاص بتخصص العمارة؛ وتوفير تكبيرات التفاصيل اللازمة لغرف محددة (تكبيرات غرف الضيوف والوحدات السكنية، وتفاصيل إنشاء الأسطح والجدران الخارجية)؛ وأبعاد ومتطلبات أداء الأبواب والنوافذ الخارجية (الأداء البصري والحراري)؛ وتحديد التخصصات المستخدمة على الأعمدة المخصصة. التخصصات المستلمة: تخطيط الموقع، الإنشاءات، السباكة، التدفئة والتهوية، الكهرباء، الأنظمة الذكية، الاقتصاد",
        "it": "La disciplina architettonica presenta file elettronici completi come base per il disegno formale della progettazione preliminare di tutte le discipline. Contenuto (sulla base della prima versione degli elaborati di lavoro architettonici): riflettere in modo globale e completo i documenti presentati da tutte le discipline e i risultati di coordinamento della fase 2; completare gli elaborati di compartimentazione antincendio, gli elaborati di difesa civile e la disposizione e le pratiche delle aperture, tutti i nomi dei locali e la quotatura propria della disciplina architettonica; fornire gli ingrandimenti di dettaglio necessari per locali specifici (ingrandimenti di camere per ospiti e unità residenziali, dettagli costruttivi di coperture e pareti esterne); le dimensioni e i requisiti prestazionali di porte e finestre esterne (prestazioni ottiche e termiche); annotare le discipline utilizzatrici sui vani riservati. Discipline destinatarie: pianificazione del sito, strutture, idrici, HVAC, elettrici, sistemi intelligenti, economia",
        "vi": "Chuyên ngành kiến trúc nộp bộ hồ sơ điện tử hoàn chỉnh làm bản nền để các chuyên ngành vẽ chính thức thiết kế cơ sở. Nội dung (trên cơ sở bản vẽ làm việc kiến trúc bản 1): phản ánh toàn diện, đầy đủ tài liệu các chuyên ngành nêu và kết quả phối hợp giai đoạn 2; hoàn thiện bản vẽ phân vùng chống cháy, mặt bằng phòng hộ dân sự và bố trí, cấu tạo cửa mở, toàn bộ tên phòng, ghi kích thước của bản thân chuyên ngành kiến trúc; cung cấp bản vẽ chi tiết phóng to cần thiết cho phòng đặc thù (phóng to phòng khách, đơn nguyên nhà ở, chi tiết cấu tạo mái và tường ngoài); kích thước và yêu cầu tính năng cửa đi, cửa sổ ngoài (tính năng quang học, nhiệt); ghi chú chuyên ngành sử dụng trên giếng kỹ thuật dự trữ. Các chuyên ngành nhận: tổng mặt bằng, kết cấu, cấp thoát nước, điều hòa, điện, hệ thống thông minh, kinh tế",
        "th": "สาขาสถาปัตยกรรมยื่นไฟล์อิเล็กทรอนิกส์ฉบับสมบูรณ์เพื่อใช้เป็นแบบร่างสำหรับการเขียนแบบทางการของการออกแบบขั้นต้นของทุกสาขา เนื้อหา (บนพื้นฐานแบบทำงานสถาปัตยกรรมฉบับที่ 1): สะท้อนเอกสารที่ทุกสาขาเสนอและผลการประสานงานของระยะที่ 2 อย่างครบถ้วนสมบูรณ์ ปรับปรุงแบบแบ่งเขตอัคคีภัย แบบป้องกันภัยทางอากาศและการจัดวางและวิธีการช่องเปิด ชื่อห้องทั้งหมด และการกำหนดขนาดของสาขาสถาปัตยกรรมเอง ให้ภาพขยายรายละเอียดที่จำเป็นสำหรับห้องเฉพาะ (ภาพขยายห้องพักและชุดที่อยู่อาศัย รายละเอียดการก่อสร้างหลังคาและผนังภายนอก) ขนาดและข้อกำหนดสมรรถนะของประตูหน้าต่างภายนอก (สมรรถนะเชิงแสงและเชิงความร้อน) ระบุสาขาผู้ใช้บนปล่องที่สำรองไว้ สาขาที่รับ: การวางผังบริเวณ โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ เศรษฐศาสตร์",
    },
    "总图专业提出，作为各专业初步设计正式绘图用底图。内容（在第一阶段基础上）：含规划用地坐标图在内的各建筑单体轴线或外墙定位图；各单体建筑±0.000标高与绝对标高的关系；红线内已有建筑、构筑物、道路、绿化等布置情况及名称；市政管线走向及平面位置；其他影响设计的资料（排洪沟、保留市政管线等）。接受专业：建筑、结构、水、暖、电、经济": {
        "en": "Proposed by the site planning discipline as the base drawings for the formal drafting of the preliminary design by all disciplines. Contents (on the basis of Stage 1): the axis or external wall location drawings of each building unit including the planned land coordinate map; the relationship between the ±0.000 elevation of each building unit and the absolute elevation; the layout and names of existing buildings, structures, roads, green areas, etc. within the red line; the alignment and plan positions of municipal pipelines; other materials affecting the design (flood drainage channels, retained municipal pipelines, etc.). Receiving disciplines: architectural, structural, plumbing, HVAC, electrical, economics",
        "ja": "総図専門が提出し、各専門の基本設計正式製図用の下図とする。内容（第一段階を基礎とする）：計画用地座標図を含む各建築単体の軸線または外壁位置図；各単体建築の±0.000標高と絶対標高の関係；赤線内の既存建築・構築物・道路・緑化等の配置状況と名称；市政配管の走向と平面位置；設計に影響するその他の資料（排洪溝、保留市政配管等）。受取専門：建築、構造、水、暖、電、経済",
        "ko": "총도 전문이 제출하며 각 전문의 기본설계 정식 도면 작성을 위한 밑그림으로 사용합니다. 내용(제1단계를 기반으로): 계획 용지 좌표도가 포함된 각 건축 단일체의 축선 또는 외벽 위치도; 각 단일 건축의 ±0.000 표고와 절대 표고의 관계; 경계선 내 기존 건축·구조물·도로·녹화 등의 배치 상황과 명칭; 시정 배관의 방향과 평면 위치; 설계에 영향을 주는 기타 자료(배수로, 보존 시정 배관 등). 수신 전문: 건축, 구조, 급배수, 난방, 전기, 경제",
        "fr": "Présenté par la discipline d'aménagement du site comme fonds de plan pour le dessin formel du projet d'esquisse de toutes les disciplines. Contenu (sur la base de la phase 1) : les plans de localisation des axes ou des murs extérieurs de chaque bâtiment, y compris le plan de coordonnées du terrain ; la relation entre la cote ±0.000 de chaque bâtiment et l'altitude absolue ; la disposition et les noms des bâtiments, ouvrages, voiries, espaces verts existants dans la ligne rouge ; le tracé et les positions en plan des réseaux municipaux ; les autres documents influant sur la conception (fossés d'évacuation des crues, réseaux municipaux conservés, etc.). Disciplines destinataires : architecture, structure, plomberie, CVC, électricité, économie",
        "de": "Vorgelegt von der Disziplin Standortplanung als Vorlagen für die formelle Zeichnung der Vorplanung aller Fachdisziplinen. Inhalt (auf der Grundlage von Stufe 1): die Achsen- oder Außenwand-Lagepläne jedes Gebäudes einschließlich des Koordinatenplans des geplanten Grundstücks; das Verhältnis der ±0.000-Höhe jedes Gebäudes zur absoluten Höhe; die Anordnung und Namen der bestehenden Gebäude, Bauwerke, Straßen, Grünflächen usw. innerhalb der roten Linie; der Verlauf und die Grundrisslagen der kommunalen Leitungen; sonstige entwurfsrelevante Unterlagen (Hochwasserentlastungsgräben, erhaltene kommunale Leitungen usw.). Empfangende Disziplinen: Architektur, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de urbanización como base para el dibujo formal del diseño preliminar de todas las especialidades. Contenido (sobre la base de la fase 1): los planos de localización de ejes o muros exteriores de cada edificio, incluido el plano de coordenadas del terreno; la relación entre la cota ±0.000 de cada edificio y la altitud absoluta; la disposición y los nombres de los edificios, obras, viales y zonas verdes existentes dentro de la línea roja; el trazado y las posiciones en planta de las redes municipales; los demás documentos que influyen en el diseño (canales de evacuación de avenidas, redes municipales conservadas, etc.). Especialidades receptoras: arquitectura, estructuras, fontanería, climatización, electricidad, economía",
        "pt": "Apresentado pela especialidade de urbanismo como base para o desenho formal do projeto preliminar de todas as especialidades. Conteúdo (com base na fase 1): os desenhos de localização dos eixos ou paredes externas de cada edifício, incluindo a planta de coordenadas do terreno; a relação entre a cota ±0.000 de cada edifício e a altitude absoluta; a disposição e os nomes dos edifícios, obras, vias e áreas verdes existentes dentro da linha de limite; o traçado e as posições em planta das redes municipais; os demais documentos que influenciam o projeto (canais de escoamento de cheias, redes municipais preservadas, etc.). Especialidades receptoras: arquitetura, estruturas, hidráulica, climatização, elétrica, economia",
        "ru": "Представляется специальностью генплана в качестве подосновы для официального оформления чертежей стадии «проект» всеми специальностями. Содержание (на основе 1-го этапа): планы расположения осей или наружных стен каждого здания, включая координатный план земельного участка; соотношение отметки ±0.000 каждого здания с абсолютной отметкой; размещение и наименования существующих зданий, сооружений, дорог, озеленения и т.п. в пределах красных линий; трассировка и плановое расположение городских сетей; прочие материалы, влияющие на проектирование (водоотводные каналы, сохраняемые городские сети и т.п.). Принимающие специальности: архитектура, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, экономика",
        "ar": "يقدمه تخصص تخطيط الموقع كأساس للرسم الرسمي للتصميم الأولي لجميع التخصصات. المحتوى (استناداً إلى المرحلة الأولى): مخططات تحديد محاور أو جدران خارجية لكل مبنى، بما في ذلك مخطط إحداثيات الأرض المخطط لها؛ والعلاقة بين منسوب ±0.000 لكل مبنى والمنسوب المطلق؛ وترتيب وأسماء المباني والمنشآت والطرق والمساحات الخضراء القائمة داخل الخط الأحمر؛ واتجاهات ومواقع الشبكات البلدية في المخطط؛ والمواد الأخرى المؤثرة في التصميم (قنوات تصريف السيول والشبكات البلدية المحفوظة وغيرها). التخصصات المستلمة: العمارة، الإنشاءات، السباكة، التدفئة والتهوية، الكهرباء، الاقتصاد",
        "it": "Presentato dalla disciplina di pianificazione del sito come base per il disegno formale della progettazione preliminare di tutte le discipline. Contenuto (sulla base della fase 1): gli elaborati di localizzazione degli assi o delle pareti esterne di ciascun edificio, inclusa la planimetria con coordinate del terreno; il rapporto tra la quota ±0.000 di ciascun edificio e la quota assoluta; la disposizione e i nomi di edifici, opere, strade, aree verdi esistenti entro la linea rossa; l'andamento e le posizioni in pianta delle reti comunali; gli altri documenti che influenzano la progettazione (canali di smaltimento piene, reti comunali conservate, ecc.). Discipline destinatarie: architettura, strutture, idrici, HVAC, elettrici, economia",
        "vi": "Do chuyên ngành tổng mặt bằng đưa ra, dùng làm bản nền để các chuyên ngành vẽ chính thức thiết kế cơ sở. Nội dung (trên cơ sở giai đoạn 1): bản vẽ định vị trục hoặc tường ngoài của từng khối nhà, kể cả bản đồ tọa độ đất quy hoạch; quan hệ giữa cao độ ±0.000 của từng khối nhà với cao độ tuyệt đối; bố trí và tên gọi nhà, công trình, đường, cây xanh hiện có trong ranh giới đỏ; hướng tuyến và vị trí mặt bằng đường ống hạ tầng; các tài liệu khác ảnh hưởng đến thiết kế (mương xả lũ, đường ống hạ tầng giữ lại, v.v.). Các chuyên ngành nhận: kiến trúc, kết cấu, cấp thoát nước, điều hòa, điện, kinh tế",
        "th": "เสนอโดยสาขาการวางผังบริเวณ ใช้เป็นแบบร่างสำหรับการเขียนแบบทางการของการออกแบบขั้นต้นของทุกสาขา เนื้อหา (บนพื้นฐานระยะที่ 1): แบบกำหนดตำแหน่งแกนหรือผนังภายนอกของอาคารแต่ละหน่วย รวมถึงแบบพิกัดที่ดินวางแผน ความสัมพันธ์ระหว่างระดับ ±0.000 ของอาคารแต่ละหน่วยกับระดับสัมบูรณ์ การจัดวางและชื่อของอาคาร สิ่งก่อสร้าง ถนน พื้นที่เขียวที่มีอยู่ภายในเส้นเขตแดง ทิศทางและตำแหน่งในผังของท่อสาธารณูปโภค เอกสารอื่นที่กระทบการออกแบบ (ร่องระบายน้ำหลาก ท่อสาธารณูปโภคที่สงวนไว้ ฯลฯ) สาขาที่รับ: สถาปัตยกรรม โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า เศรษฐศาสตร์",
    },
    "此时应提清所有相关的配合资料：水专业向电专业提出各层消火栓、水流指示器的数量；暖专业向电专业提出各层防火阀、排烟阀、加压送风阀的数量；完善、补充和确认第二阶段向各相关专业提出的资料；水、电专业向智能化专业提出需远程计量各类表具的数量、污水池污水泵的数量；暖专业向智能化专业提出空调通风设备自动化控制要求": {
        "en": "At this time all relevant coordination materials shall be clarified and submitted: the plumbing discipline raises to the electrical discipline the quantities of fire hydrants and water flow indicators on each floor; the HVAC discipline raises to the electrical discipline the quantities of fire dampers, smoke exhaust valves and pressurization air supply valves on each floor; improve, supplement and confirm the materials raised to all relevant disciplines in Stage 2; the plumbing and electrical disciplines raise to the intelligent systems discipline the quantities of various meters requiring remote metering and the quantity of sewage pit pumps; the HVAC discipline raises to the intelligent systems discipline the automatic control requirements for air-conditioning and ventilation equipment",
        "ja": "この時点で関連する全ての配合資料を明確化して提出する：水専門から電専門へ各階の消火栓・水流指示器の数量を提出する；暖専門から電専門へ各階の防火ダンパー・排煙弁・加圧給気弁の数量を提出する；第二段階で各関連専門へ提出した資料を改善・補充・確認する；水・電専門から智能化専門へ遠隔計量が必要な各種メーター類の数量・汚水池汚水ポンプの数量を提出する；暖専門から智能化専門へ空調換気設備の自動制御要求を提出する",
        "ko": "이때 관련된 모든 배합 자료를 명확히 정리하여 제출합니다: 급배수 전문이 전기 전문에게 각 층 소화전·수류지시기의 수량을 제출합니다; 난방 전문이 전기 전문에게 각 층 방화댐퍼·제연밸브·가압송풍밸브의 수량을 제출합니다; 제2단계에서 각 관련 전문에게 제출한 자료를 개선·보완·확인합니다; 급배수·전기 전문이 지능화 전문에게 원격 계량이 필요한 각종 계기류의 수량·오수조 오수펌프의 수량을 제출합니다; 난방 전문이 지능화 전문에게 공조 환기 설비 자동 제어 요구를 제출합니다",
        "fr": "À ce stade, tous les documents de coordination pertinents doivent être clarifiés et soumis : la plomberie soumet à l'électricité les quantités de poteaux d'incendie et d'indicateurs de débit d'eau à chaque niveau ; le CVC soumet à l'électricité les quantités de clapets coupe-feu, de clapets de désenfumage et de clapets de soufflage de pressurisation à chaque niveau ; améliorer, compléter et confirmer les documents soumis à toutes les disciplines concernées en phase 2 ; la plomberie et l'électricité soumettent aux systèmes intelligents les quantités des divers compteurs nécessitant une télérélève et la quantité des pompes de puisards d'eaux usées ; le CVC soumet aux systèmes intelligents les exigences de commande automatique des équipements de CVC et de ventilation",
        "de": "Zu diesem Zeitpunkt sind alle relevanten Koordinationsunterlagen klarzustellen und einzureichen: die Sanitärdisziplin meldet der Elektrodisziplin die Anzahl der Hydranten und Wasserströmungswächter je Geschoss; die Heizung-Lüftungs-Disziplin meldet der Elektrodisziplin die Anzahl der Brandschutzklappen, Rauchabzugsklappen und Druckbelüftungsklappen je Geschoss; die in Stufe 2 an alle betroffenen Disziplinen eingereichten Unterlagen verbessern, ergänzen und bestätigen; die Sanitär- und Elektrodisziplin melden der Gebäudeautomations-Disziplin die Anzahl der fernauslesbaren Messgeräte und die Anzahl der Abwasserhebeanlagen in Abwassergruben; die Heizung-Lüftungs-Disziplin meldet der Gebäudeautomations-Disziplin die Anforderungen an die automatische Steuerung der Klima- und Lüftungsgeräte",
        "es": "En este momento deben aclararse y presentarse todos los documentos de coordinación pertinentes: fontanería presenta a electricidad las cantidades de hidrantes e indicadores de flujo de agua en cada nivel; climatización presenta a electricidad las cantidades de compuertas cortafuegos, válvulas de extracción de humos y válvulas de impulsión de presurización en cada nivel; mejorar, complementar y confirmar los documentos presentados a todas las especialidades afectadas en la fase 2; fontanería y electricidad presentan a sistemas inteligentes las cantidades de los diversos contadores que requieren telelectura y la cantidad de bombas de sumideros de aguas residuales; climatización presenta a sistemas inteligentes los requisitos de control automático de los equipos de climatización y ventilación",
        "pt": "Neste momento devem ser esclarecidos e apresentados todos os documentos de coordenação pertinentes: hidráulica apresenta à elétrica as quantidades de hidrantes e indicadores de fluxo de água em cada pavimento; climatização apresenta à elétrica as quantidades de dampers corta-fogo, válvulas de exaustão de fumaça e válvulas de insuflação de pressurização em cada pavimento; melhorar, complementar e confirmar os documentos apresentados a todas as especialidades envolvidas na fase 2; hidráulica e elétrica apresentam aos sistemas inteligentes as quantidades dos diversos medidores que exigem telemetria e a quantidade de bombas de poços de esgoto; climatização apresenta aos sistemas inteligentes os requisitos de controle automático dos equipamentos de climatização e ventilação",
        "ru": "На данном этапе следует окончательно представить все соответствующие материалы профессиональной координации: водоснабжение представляет электроснабжению количество пожарных гидрантов и указателей потока воды на каждом этаже; отопление и вентиляция представляет электроснабжению количество противопожарных клапанов, клапанов дымоудаления и клапанов подпора воздуха на каждом этаже; улучшить, дополнить и подтвердить материалы, представленные на 2-м этапе всем соответствующим специальностям; водоснабжение и электроснабжение представляют интеллектуальным системам количество приборов учёта, требующих дистанционного снятия показаний, и количество насосов в приёмных резервуарах сточных вод; отопление и вентиляция представляет интеллектуальным системам требования к автоматическому управлению оборудованием кондиционирования и вентиляции",
        "ar": "في هذه المرحلة يجب توضيح وتقديم جميع مواد التنسيق ذات الصلة: تقدم السباكة إلى الكهرباء أعداد صنابير الإطفاء ومؤشرات تدفق المياه في كل طابق؛ ويقدم التدفئة والتهوية إلى الكهرباء أعداد مخمدات الحريق وصمامات إخراج الدخان وصمامات الهواء المضغوط في كل طابق؛ وتحسين واستكمال وتأكيد المواد المقدمة إلى جميع التخصصات المعنية في المرحلة الثانية؛ وتقدم السباكة والكهرباء إلى الأنظمة الذكية أعداد أجهزة القياس المختلفة التي تتطلب قياساً عن بعد وأعداد مضخات أحواض الصرف الصحي؛ ويقدم التدفئة والتهوية إلى الأنظمة الذكية متطلبات التحكم الآلي لمعدات التكييف والتهوية",
        "it": "In questo momento devono essere chiariti e presentati tutti i documenti di coordinamento pertinenti: gli idrici presentano agli elettrici le quantità di idranti e indicatori di flusso d'acqua a ogni piano; l'HVAC presenta agli elettrici le quantità di serrande tagliafuoco, valvole di estrazione fumi e valvole di immissione pressurizzata a ogni piano; migliorare, integrare e confermare i documenti presentati a tutte le discipline interessate nella fase 2; idrici ed elettrici presentano ai sistemi intelligenti le quantità dei vari contatori che richiedono telelettura e la quantità di pompe dei pozzetti di scarico; l'HVAC presenta ai sistemi intelligenti i requisiti di controllo automatico delle apparecchiature di climatizzazione e ventilazione",
        "vi": "Lúc này cần làm rõ và nộp toàn bộ tài liệu phối hợp liên quan: cấp thoát nước nêu cho điện số lượng họng cứu hỏa và chỉ thị lưu lượng nước từng tầng; điều hòa nêu cho điện số lượng van chống cháy, van xả khói, van cấp gió tăng áp từng tầng; hoàn thiện, bổ sung và xác nhận tài liệu đã nêu cho các chuyên ngành liên quan ở giai đoạn 2; cấp thoát nước và điện nêu cho hệ thống thông minh số lượng các loại đồng hồ cần đo từ xa và số lượng bơm hố nước thải; điều hòa nêu cho hệ thống thông minh yêu cầu điều khiển tự động thiết bị điều hòa thông gió",
        "th": "ในเวลานี้ต้องแจ้งเอกสารการประสานงานที่เกี่ยวข้องทั้งหมดให้ชัดเจน: ประปาเสนอต่อไฟฟ้าถึงจำนวนหัวจ่ายน้ำดับเพลิงและตัวชี้วัดการไหลของน้ำในแต่ละชั้น ระบบปรับอากาศเสนอต่อไฟฟ้าถึงจำนวนลิ้นกันไฟ ลิ้นระบายควัน และลิ้นจ่ายอากาศอัดในแต่ละชั้น ปรับปรุง เติมเต็ม และยืนยันเอกสารที่เสนอต่อทุกสาขาที่เกี่ยวข้องในระยะที่ 2 ประปาและไฟฟ้าเสนอต่อระบบอัจฉริยะถึงจำนวนมิเตอร์ชนิดต่าง ๆ ที่ต้องวัดระยะไกลและจำนวนปั๊มบ่อน้ำเสีย ระบบปรับอากาศเสนอต่อระบบอัจฉริยะถึงข้อกำหนดการควบคุมอัตโนมัติของอุปกรณ์ปรับอากาศและระบายอากาศ",
    },
    "在第二版建筑作业图和机电专业确认与提清资料并协调后的基础上、由于管道综合后发生的局部变更（重大变更应在第二版作业图中完成），采用纸质文件形式将局部变更处提交给相关专业": {
        "en": "On the basis of the second version of the architectural working drawings and the confirmed and clarified materials of the mechanical and electrical disciplines after coordination, for the local changes resulting from pipe comprehensive coordination (major changes shall be completed in the second version of the working drawings), the local changes shall be submitted to the relevant disciplines in paper document form",
        "ja": "第二版建築作業図と電気・機械専門の確認・提清資料を調整した上での基礎の上で、配管総合調整後に生じた局部変更（重大な変更は第二版作業図で完了させる）について、紙文書の形で局部変更箇所を関連専門に提出する",
        "ko": "제2판 건축 작업도와 전기·기계 전문의 확인·정리 자료를 조정한 기초 위에서, 배관 종합 후 발생한 국부 변경(중대 변경은 제2판 작업도에서 완료)은 종이 문서 형태로 국부 변경 지점을 관련 전문에게 제출합니다",
        "fr": "Sur la base de la deuxième version des plans de travail d'architecture et des documents confirmés et clarifiés des disciplines électriques et mécaniques après coordination, pour les modifications locales résultant de la coordination complète des réseaux (les modifications majeures doivent être achevées dans la deuxième version des plans de travail), les modifications locales doivent être soumises aux disciplines concernées sous forme de documents papier",
        "de": "Auf der Grundlage der zweiten Fassung der Architektur-Arbeitszeichnungen und der nach Koordination bestätigten und klargestellten Unterlagen der elektro- und maschinentechnischen Disziplinen sind die sich aus der Rohrleitungsgesamtkoordination ergebenden lokalen Änderungen (wesentliche Änderungen sind in der zweiten Fassung der Arbeitszeichnungen abzuschließen) den betroffenen Disziplinen in Papierform zu übermitteln",
        "es": "Sobre la base de la segunda versión de los planos de trabajo de arquitectura y de los documentos confirmados y aclarados de las especialidades eléctrica y mecánica tras la coordinación, las modificaciones locales derivadas de la coordinación integral de redes (las modificaciones importantes deben completarse en la segunda versión de los planos de trabajo) deben presentarse a las especialidades afectadas en forma de documentos en papel",
        "pt": "Com base na segunda versão dos desenhos de trabalho de arquitetura e nos documentos confirmados e esclarecidos das especialidades elétrica e mecânica após coordenação, as alterações locais resultantes da coordenação integral de tubulações (alterações importantes devem ser concluídas na segunda versão dos desenhos de trabalho) devem ser apresentadas às especialidades envolvidas em forma de documentos impressos",
        "ru": "На основе второй версии рабочих архитектурных чертежей и подтверждённых и окончательно представленных материалов электротехнической и механической специальностей после координации, локальные изменения, возникшие в результате комплексной увязки инженерных сетей (значительные изменения должны быть завершены во второй версии рабочих чертежей), представляются соответствующим специальностям в форме бумажных документов",
        "ar": "استناداً إلى النسخة الثانية من مخططات العمل المعمارية والمواد المؤكدة والموضحة لتخصصي الكهرباء والميكانيكا بعد التنسيق، يجب تقديم التعديلات المحلية الناتجة عن التنسيق الشامل للأنابيب (يجب إكمال التعديلات الهامة في النسخة الثانية من مخططات العمل) إلى التخصصات المعنية في شكل مستندات ورقية",
        "it": "Sulla base della seconda versione degli elaborati di lavoro architettonici e dei documenti confermati e chiariti delle discipline elettrica e meccanica dopo il coordinamento, le modifiche locali derivanti dal coordinamento integrale delle tubazioni (le modifiche importanti devono essere completate nella seconda versione degli elaborati di lavoro) devono essere presentate alle discipline interessate in forma di documenti cartacei",
        "vi": "Trên cơ sở bản vẽ làm việc kiến trúc bản 2 và tài liệu đã xác nhận, làm rõ của chuyên ngành điện, cơ khí sau khi phối hợp, các thay đổi cục bộ phát sinh sau tổng hợp đường ống (thay đổi lớn phải hoàn thành trong bản vẽ làm việc bản 2) được nộp cho các chuyên ngành liên quan dưới dạng hồ sơ giấy",
        "th": "บนพื้นฐานของแบบทำงานสถาปัตยกรรมฉบับที่ 2 และเอกสารที่ยืนยันและชี้แจงของสาขาไฟฟ้าและเครื่องกลหลังการประสานงาน การเปลี่ยนแปลงเฉพาะจุดที่เกิดจากการประสานงานท่อโดยรวม (การเปลี่ยนแปลงสำคัญต้องเสร็จในแบบทำงานฉบับที่ 2) ให้ส่งจุดที่เปลี่ยนแปลงเฉพาะจุดให้สาขาที่เกี่ยวข้องในรูปแบบเอกสารกระดาษ",
    },
    "项目设计主持人组织、各专业负责人参加（必要时设计人参加）。工作内容：验证与核对房间名称和使用功能；验证与核对防火分区；对主要管道进行综合与协调以保证不给施工图设计留下后患；验证与核对各专业所提资料和配合过程中的协商内容是否在相关图纸中得到反映；验证是否满足甲方和政府主管部门要求。管道综合后预留洞位置和尺寸若发生变化，各专业应立即书面通知结构专业（本阶段 78%P～87%P 时间周期内完成）": {
        "en": "Organized by the chief designer of the project, attended by the person in charge of each discipline (designers participate when necessary). Work contents: verify and check room names and usage functions; verify and check fire compartmentation; comprehensively coordinate the main pipelines to ensure that no problems are left for the construction document design; verify and check whether the materials submitted by each discipline and the negotiated contents in the coordination process are reflected in the relevant drawings; verify whether the requirements of the owner and government authorities are met. If the positions and dimensions of reserved openings change after the pipe comprehensive coordination, each discipline shall immediately notify the structural discipline in writing (to be completed within the 78%P~87%P time period of this stage)",
        "ja": "プロジェクト設計主催者が組織し、各専門の担当者が参加する（必要な場合設計者も参加）。作業内容：室名と使用機能を検証・照合する；防火区画を検証・照合する；主要配管を総合調整し、施工図設計に問題を残さないようにする；各専門の提出資料と配合過程での協議内容が関連図面に反映されているかを検証・照合する；甲方と政府主管部門の要求を満たしているかを検証する。配管総合後、予備開口の位置寸法が変わった場合、各専門は直ちに書面で構造専門に通知する（本段階 78%P～87%P 期間内に完了）",
        "ko": "프로젝트 설계 주관자가 조직하고 각 전문의 책임자가 참여합니다(필요 시 설계자도 참여). 작업 내용: 실명과 사용 기능을 검증·대조합니다; 방화구획을 검증·대조합니다; 주요 배관을 종합 조정하여 시공도 설계에 문제를 남기지 않도록 합니다; 각 전문이 제출한 자료와 배합 과정에서 협의한 내용이 관련 도면에 반영되었는지 검증·대조합니다; 갑과 정부 주관 부문의 요구를 충족하는지 검증합니다. 배관 종합 후 예비 구멍의 위치와 치수가 변경되면 각 전문은 즉시 서면으로 구조 전문에게 통지합니다(본 단계 78%P~87%P 기간 내 완료)",
        "fr": "Organisée par le concepteur en chef du projet, à laquelle participent les responsables de chaque discipline (les concepteurs participent si nécessaire). Contenu des travaux : vérifier et contrôler les noms des pièces et leurs fonctions d'utilisation ; vérifier et contrôler le compartimentage coupe-feu ; coordonner globalement les principales canalisations afin de ne laisser aucun problème pour la conception des documents d'exécution ; vérifier et contrôler que les documents soumis par chaque discipline et le contenu des négociations au cours du processus de coordination sont reflétés dans les dessins concernés ; vérifier que les exigences du maître d'ouvrage et des autorités gouvernementales sont satisfaites. Si les positions et dimensions des réservations changent après la coordination globale des réseaux, chaque discipline doit immédiatement notifier par écrit la discipline structure (à achever dans la période de 78 %P à 87 %P de cette phase)",
        "de": "Organisiert vom Chefplaner des Projekts, Teilnahme der Verantwortlichen jeder Disziplin (bei Bedarf nehmen auch Planer teil). Arbeitsinhalte: Raumnamen und Nutzungsfunktionen verifizieren und abgleichen; Brandabschnitte verifizieren und abgleichen; die Hauptleitungen umfassend koordinieren, um sicherzustellen, dass keine Probleme für die Ausführungsplanung zurückbleiben; verifizieren und abgleichen, ob die von jeder Disziplin eingereichten Unterlagen und die im Koordinationsprozess ausgehandelten Inhalte in den betreffenden Zeichnungen widergespiegelt sind; verifizieren, ob die Anforderungen des Bauherrn und der zuständigen Behörden erfüllt sind. Ändern sich Position und Abmessungen der Durchbrüche nach der Rohrleitungskoordination, hat jede Disziplin die Tragwerksdisziplin unverzüglich schriftlich zu benachrichtigen (innerhalb des Zeitraums 78 %P bis 87 %P dieser Stufe abzuschließen)",
        "es": "Organizada por el proyectista jefe del proyecto, con asistencia de los responsables de cada especialidad (los proyectistas participan cuando sea necesario). Contenido de los trabajos: verificar y comprobar los nombres de los locales y sus funciones de uso; verificar y comprobar la compartimentación contra incendios; coordinar integralmente las principales canalizaciones para no dejar problemas para el diseño de los documentos de obra; verificar y comprobar que los documentos presentados por cada especialidad y el contenido de las negociaciones durante el proceso de coordinación se reflejan en los planos correspondientes; verificar que se cumplen los requisitos del cliente y de las autoridades gubernamentales. Si cambian las posiciones y dimensiones de los huecos reservados tras la coordinación integral de redes, cada especialidad deberá notificar inmediatamente por escrito a la especialidad de estructuras (debe completarse dentro del período 78 %P a 87 %P de esta fase)",
        "pt": "Organizada pelo projetista-chefe do projeto, com participação dos responsáveis de cada especialidade (os projetistas participam quando necessário). Conteúdo do trabalho: verificar e conferir os nomes dos ambientes e suas funções de uso; verificar e conferir a compartimentação contra incêndio; coordenar integralmente as principais tubulações para garantir que não fiquem problemas para o projeto de execução; verificar e conferir se os documentos apresentados por cada especialidade e os conteúdos negociados durante o processo de coordenação estão refletidos nos desenhos correspondentes; verificar se os requisitos do contratante e das autoridades governamentais são atendidos. Se as posições e dimensões das aberturas reservadas mudarem após a coordenação integral das tubulações, cada especialidade deverá notificar imediatamente por escrito a especialidade de estruturas (a concluir dentro do período de 78%P a 87%P desta fase)",
        "ru": "Организуется главным проектировщиком проекта, участие руководителей каждой специальности (при необходимости — проектировщиков). Содержание работ: проверить и сверить наименования помещений и их функциональное назначение; проверить и сверить противопожарные отсеки; провести комплексную увязку основных инженерных сетей, чтобы не оставить проблем для стадии «рабочая документация»; проверить и сверить, отражены ли в соответствующих чертежах материалы, представленные каждой специальностью, и согласованные в процессе координации вопросы; проверить соблюдение требований заказчика и органов государственного надзора. Если после комплексной увязки сетей изменятся положение или размеры отверстий, каждая специальность должна немедленно в письменной форме уведомить конструктивную специальность (завершить в пределах 78 %P–87 %P периода данного этапа)",
        "ar": "ينظمها رئيس المصممين للمشروع، ويشارك فيها المسؤولون عن كل تخصص (يشارك المصممون عند الحاجة). محتوى العمل: التحقق ومطابقة أسماء الغرف ووظائف الاستخدام؛ والتحقق ومطابقة تقسيمات الحماية من الحرائق؛ والتنسيق الشامل للأنابيب الرئيسية لضمان عدم ترك مشاكل لتصميم مستندات التنفيذ؛ والتحقق ومطابقة ما إذا كانت المواد المقدمة من كل تخصص ومحتوى التفاوض أثناء عملية التنسيق تنعكس في المخططات ذات الصلة؛ والتحقق من تلبية متطلبات المالك والجهات الحكومية المختصة. إذا تغيرت مواقع وأبعاد الفتحات المخصصة بعد التنسيق الشامل للأنابيب، فيجب على كل تخصص إخطار تخصص الإنشاءات كتابياً فوراً (يجب إكماله خلال فترة 78%P إلى 87%P من هذه المرحلة)",
        "it": "Organizzata dal capo progettista del progetto, con partecipazione dei responsabili di ogni disciplina (i progettisti partecipano quando necessario). Contenuto del lavoro: verificare e controllare i nomi dei locali e le funzioni d'uso; verificare e controllare la compartimentazione antincendio; coordinare integralmente le tubazioni principali per garantire che non rimangano problemi per la progettazione esecutiva; verificare e controllare che i documenti presentati da ogni disciplina e i contenuti negoziati durante il processo di coordinamento siano riflessi negli elaborati pertinenti; verificare che siano soddisfatti i requisiti del committente e delle autorità governative. Se le posizioni e le dimensioni delle aperture riservate cambiano dopo il coordinamento integrale delle tubazioni, ogni disciplina deve notificare immediatamente per iscritto la disciplina strutturale (da completare entro il periodo 78%P~87%P di questa fase)",
        "vi": "Do chủ trì thiết kế dự án tổ chức, trưởng các chuyên ngành tham gia (khi cần thiết có người thiết kế tham gia). Nội dung công việc: kiểm tra, đối chiếu tên phòng và chức năng sử dụng; kiểm tra, đối chiếu phân vùng chống cháy; tổng hợp và phối hợp các đường ống chính để không để lại hậu quả cho thiết kế thi công; kiểm tra, đối chiếu xem tài liệu các chuyên ngành nêu và nội dung thỏa thuận trong quá trình phối hợp có được phản ánh trong bản vẽ liên quan không; kiểm tra xem có đáp ứng yêu cầu của chủ đầu tư và cơ quan quản lý nhà nước không. Nếu vị trí, kích thước lỗ chờ thay đổi sau khi tổng hợp đường ống, các chuyên ngành phải thông báo ngay bằng văn bản cho chuyên ngành kết cấu (hoàn thành trong khoảng 78%P~87%P thời gian giai đoạn này)",
        "th": "จัดโดยหัวหน้าผู้ออกแบบโครงการ ผู้รับผิดชอบแต่ละสาขาเข้าร่วม (ผู้ออกแบบเข้าร่วมเมื่อจำเป็น) เนื้อหางาน: ตรวจสอบและทวนสอบชื่อห้องและหน้าที่การใช้งาน ตรวจสอบและทวนสอบการแบ่งเขตอัคคีภัย ประสานงานท่อหลักโดยรวมเพื่อไม่ให้เหลือปัญหาไว้สำหรับการออกแบบแบบก่อสร้าง ตรวจสอบและทวนสอบว่าเอกสารที่แต่ละสาขาเสนอและเนื้อหาที่ตกลงกันระหว่างการประสานงานสะท้อนอยู่ในแบบที่เกี่ยวข้องหรือไม่ ตรวจสอบว่าเป็นไปตามข้อกำหนดของเจ้าของโครงการและหน่วยงานรัฐหรือไม่ หากตำแหน่งและขนาดของช่องเผื่อเปลี่ยนแปลงหลังการประสานงานท่อโดยรวม แต่ละสาขาต้องแจ้งสาขาโครงสร้างเป็นลายลักษณ์อักษรทันที (ให้เสร็จภายในช่วง 78%P ถึง 87%P ของระยะนี้)",
    },
    "针对第二阶段水、暖、电专业对总图所提资料（各市政管线接口位置、室外管线布置、室外构筑物位置等）进行确认；根据管道综合情况进行局部调整与变更": {
        "en": "Confirm the materials raised by the plumbing, HVAC and electrical disciplines to the site planning discipline in Stage 2 (interface positions of various municipal pipelines, outdoor pipeline layouts, positions of outdoor structures, etc.); make local adjustments and changes according to the pipe comprehensive coordination situation",
        "ja": "第二段階で水・暖・電専門が総図専門へ提出した資料（各市政配管の接続位置、屋外配管の配置、屋外構築物の位置等）を確認する；配管総合の状況に応じて局部調整と変更を行う",
        "ko": "제2단계에서 급배수·난방·전기 전문이 총도 전문에게 제출한 자료(각 시정 배관 접속 위치, 옥외 배관 배치, 옥외 구조물 위치 등)를 확인합니다; 배관 종합 상황에 따라 국부 조정과 변경을 실시합니다",
        "fr": "Confirmer les documents soumis par les disciplines plomberie, CVC et électricité à l'aménagement du site en phase 2 (positions des raccordements des divers réseaux municipaux, dispositions des réseaux extérieurs, positions des ouvrages extérieurs, etc.) ; procéder à des ajustements et modifications locaux en fonction de la situation de coordination globale des réseaux",
        "de": "Die in Stufe 2 von den Disziplinen Sanitär, Heizung-Lüftung und Elektro an die Standortplanung eingereichten Unterlagen (Anschlusspositionen der städtischen Leitungen, Außenleitungstrassen, Positionen von Außenbauwerken usw.) bestätigen; entsprechend der Rohrleitungskoordination lokale Anpassungen und Änderungen vornehmen",
        "es": "Confirmar los documentos presentados por las especialidades de fontanería, climatización y electricidad a la urbanización en la fase 2 (posiciones de conexión de los distintos redes municipales, trazados de redes exteriores, posiciones de obras exteriores, etc.); realizar ajustes y modificaciones locales según la situación de la coordinación integral de redes",
        "pt": "Confirmar os documentos apresentados pelas especialidades de hidráulica, climatização e elétrica ao urbanismo na fase 2 (posições de conexão das diversas redes municipais, traçados de redes externas, posições de obras externas, etc.); realizar ajustes e alterações locais conforme a situação da coordenação integral das tubulações",
        "ru": "Подтвердить материалы, представленные на 2-м этапе специальностями водоснабжения, отопления и вентиляции и электроснабжения специальности генплана (места присоединения городских сетей, наружные сети, расположение наружных сооружений и т.п.) ; вносить локальные корректировки и изменения по результатам комплексной увязки сетей",
        "ar": "تأكيد المواد المقدمة من تخصصات السباكة والتدفئة والتهوية والكهرباء إلى تخطيط الموقع في المرحلة الثانية (مواقع توصيل الشبكات البلدية المختلفة، وتخطيطات الشبكات الخارجية، ومواقع المنشآت الخارجية وغيرها)؛ وإجراء تعديلات وتغييرات محلية وفقاً لحالة التنسيق الشامل للأنابيب",
        "it": "Confermare i documenti presentati dalle discipline idriche, HVAC ed elettriche alla pianificazione del sito nella fase 2 (posizioni dei collegamenti delle varie reti comunali, disposizioni delle reti esterne, posizioni delle opere esterne, ecc.); apportare adeguamenti e modifiche locali in base alla situazione del coordinamento integrale delle tubazioni",
        "vi": "Xác nhận tài liệu do cấp thoát nước, điều hòa, điện nêu cho tổng mặt bằng ở giai đoạn 2 (vị trí đấu nối các tuyến hạ tầng, bố trí đường ống ngoài nhà, vị trí công trình ngoài trời, v.v.); căn cứ tình hình tổng hợp đường ống để điều chỉnh và thay đổi cục bộ",
        "th": "ยืนยันเอกสารที่ประปา ระบบปรับอากาศ และไฟฟ้าเสนอต่อการวางผังบริเวณในระยะที่ 2 (ตำแหน่งเชื่อมต่อท่อสาธารณูปโภคต่าง ๆ การวางท่อภายนอก ตำแหน่งสิ่งก่อสร้างภายนอก ฯลฯ) ปรับและเปลี่ยนแปลงเฉพาะจุดตามสถานการณ์การประสานงานท่อโดยรวม",
    },
    "在会签之前，各专业应再次进行专业间核对，确保上述各阶段的结论和成果已经得到完整落实；各专业的设计图纸为已经过规定的校、审和修改程序后的图纸（设计文件归档之前完成）": {
        "en": "Before joint signing, each discipline shall re-verify among disciplines to ensure that the conclusions and results of the above stages have been fully implemented; the design drawings of each discipline shall have passed the prescribed procedures of checking, reviewing and revising (to be completed before the design documents are archived)",
        "ja": "会签前に、各専門は再び専門間の照合を行い、上記各段階の結論と成果が完全に実施されていることを確保する；各専門の設計図面は規定の校・審・修正手順を経た図面であること（設計文書のアーカイブ前に完了）",
        "ko": "회서 전에 각 전문은 다시 전문 간 대조를 수행하여 상기 각 단계의 결론과 성과가 완전히 이행되었음을 보장한다；각 전문의 설계 도면은 규정된 검교·심사·수정 절차를 거친 도면이어야 한다(설계 문서 보관 전 완료)",
        "fr": "Avant la signature conjointe, chaque spécialité doit procéder à nouveau à une vérification interdisciplinaire afin de garantir que les conclusions et les résultats des étapes ci-dessus sont pleinement mis en œuvre ; les plans de conception de chaque spécialité doivent avoir passé les procédures prescrites de vérification, de révision et de modification (à terminer avant l'archivage des documents de conception)",
        "de": "Vor der gemeinsamen Unterzeichnung muss jede Fachrichtung erneut fachübergreifende Abstimmungen vornehmen, um sicherzustellen, dass die Schlussfolgerungen und Ergebnisse der oben genannten Phasen vollständig umgesetzt wurden; die Entwurfspläne jeder Fachrichtung müssen die vorgeschriebenen Prüfungs-, Revisions- und Änderungsverfahren durchlaufen haben (vor der Archivierung der Entwurfsunterlagen abzuschließen)",
        "es": "Antes de la firma conjunta, cada especialidad debe volver a realizar la verificación interdisciplinaria para garantizar que las conclusiones y resultados de las etapas anteriores se han implementado plenamente; los planos de diseño de cada especialidad deben haber pasado los procedimientos prescritos de verificación, revisión y modificación (debe completarse antes del archivo de los documentos de diseño)",
        "pt": "Antes da assinatura conjunta, cada especialidade deve realizar novamente a verificação interdisciplinar para garantir que as conclusões e resultados das etapas acima foram plenamente implementados; os desenhos de projeto de cada especialidade devem ter passado pelos procedimentos prescritos de verificação, revisão e modificação (a concluir antes do arquivamento dos documentos de projeto)",
        "ru": "Перед совместным подписанием каждая специальность должна повторно провести межспециальную сверку, чтобы убедиться, что выводы и результаты вышеуказанных этапов полностью реализованы; чертежи проектирования каждой специальности должны пройти установленные процедуры проверки, рецензирования и внесения изменений (завершить до сдачи проектной документации в архив)",
        "ar": "قبل التوقيع المشترك، يجب على كل تخصص إجراء مراجعة مشتركة بين التخصصات مرة أخرى للتأكد من أن استنتاجات ونتائج المراحل المذكورة أعلاه قد نُفذت بالكامل؛ يجب أن تكون رسومات التصميم لكل تخصص قد خضعت لإجراءات الفحص والمراجعة والتعديل المقررة (يُكمل قبل أرشفة مستندات التصميم)",
        "it": "Prima della firma congiunta, ogni disciplina deve effettuare nuovamente la verifica interdisciplinare per garantire che le conclusioni e i risultati delle fasi sopra indicate siano stati pienamente attuati; gli elaborati progettuali di ciascuna disciplina devono aver superato le procedure prescritte di verifica, revisione e modifica (da completare prima dell'archiviazione dei documenti di progetto)",
        "vi": "Trước khi ký chung, mỗi chuyên ngành phải tiến hành đối chiếu liên chuyên ngành lần nữa để bảo đảm các kết luận và kết quả của các giai đoạn nêu trên đã được thực hiện đầy đủ; bản vẽ thiết kế của mỗi chuyên ngành phải đã qua các quy trình kiểm tra, thẩm tra và sửa đổi quy định (hoàn thành trước khi lưu trữ hồ sơ thiết kế)",
        "th": "ก่อนการลงนามร่วมกัน แต่ละสาขาต้องดำเนินการตรวจสอบระหว่างสาขาอีกครั้งเพื่อให้แน่ใจว่าข้อสรุปและผลลัพธ์ของขั้นตอนข้างต้นได้ถูกนำไปปฏิบัติอย่างครบถ้วน แบบออกแบบของแต่ละสาขาต้องผ่านขั้นตอนการตรวจทาน ตรวจสอบ และแก้ไขตามที่กำหนดไว้ (แล้วเสร็จก่อนการจัดเก็บเอกสารการออกแบบ)",
    },
    "按照院相关文件的要求进行（设计文件交付之前完成）": {
        "en": "To be carried out in accordance with the requirements of the relevant institute documents (to be completed before the delivery of design documents)",
        "ja": "院の関連文書の要求に従って実施する（設計文書の納品前に完了）",
        "ko": "원내 관련 문서의 요구에 따라 진행한다(설계 문서 인도 전 완료)",
        "fr": "À exécuter conformément aux exigences des documents pertinents de l'institut (à terminer avant la remise des documents de conception)",
        "de": "Gemäß den Anforderungen der einschlägigen Institutsunterlagen durchzuführen (vor der Übergabe der Entwurfsunterlagen abzuschließen)",
        "es": "Debe ejecutarse de acuerdo con los requisitos de los documentos pertinentes del instituto (debe completarse antes de la entrega de los documentos de diseño)",
        "pt": "A executar em conformidade com os requisitos dos documentos pertinentes do instituto (a concluir antes da entrega dos documentos de projeto)",
        "ru": "Выполнять в соответствии с требованиями соответствующих документов института (завершить до передачи проектной документации)",
        "ar": "يُنفذ وفقًا لمتطلبات الوثائق ذات الصلة الصادرة عن المعهد (يُكمل قبل تسليم مستندات التصميم)",
        "it": "Da eseguire in conformità con i requisiti dei documenti pertinenti dell'istituto (da completare prima della consegna dei documenti di progetto)",
        "vi": "Thực hiện theo yêu cầu của các văn bản liên quan của viện (hoàn thành trước khi bàn giao hồ sơ thiết kế)",
        "th": "ให้ดำเนินการตามข้อกำหนดของเอกสารที่เกี่ยวข้องของสถาบัน (แล้วเสร็จก่อนการส่งมอบเอกสารการออกแบบ)",
    },
    "项目经理组织、设计主持人主持、各专业负责人参加（施工图设计阶段开始日）：介绍项目基本情况（未经过我院初步设计的项目同时发放《设计任务书》和《设计策划表》）；发放经审批或甲方确认的建筑初步设计图；发放各专业初步设计审批意见；研究讨论甲方及政府部门要求，制定统一设计要求和措施；各专业研究需要调研和收集的资料并尽快组织调研；限额设计提出经审批（或甲方同意）的初步设计概算；制定本工程《专业配合进度表》并签字生效": {
        "en": "Organized by the project manager, chaired by the design director, attended by the heads of all disciplines (on the start date of the construction document design stage): introduce the basic project information (for projects not having passed through our institute's preliminary design, also issue the Design Task Book and the Design Planning Sheet); issue the architectural preliminary design drawings approved or confirmed by the client; issue the preliminary design approval opinions of all disciplines; study and discuss the requirements of the client and government authorities and formulate unified design requirements and measures; each discipline studies the materials to be surveyed and collected and organizes the survey as soon as possible; value engineering proposes the approved (or client-agreed) preliminary design budget estimate; prepare the Professional Coordination Schedule for this project and sign it into effect",
        "ja": "プロジェクトマネージャーが組織し、設計主催者が司会し、各専門担当責任者が参加する（施工図設計段階の開始日）：プロジェクトの基本情報を紹介する（当院の基本設計を経ていないプロジェクトには同時に『設計任務書』と『設計計画表』を配布する）；審査済みまたは発注者が確認した建築基本設計図を配布する；各専門の基本設計審査意見を配布する；発注者および政府部門の要求を研究・検討し、統一した設計要求と措置を策定する；各専門は調査・収集すべき資料を検討し、できるだけ早く調査を組織する；限界設計は審査済み（または発注者同意）の基本設計概算を提出する；本工事の『専門配合進度表』を作成し、署名して発効させる",
        "ko": "프로젝트 매니저가 조직하고 설계 책임자가 주재하며 각 전문 담당자가 참석합니다(시공도 설계 단계 시작일): 프로젝트 기본 정보를 소개합니다(당원의 기본설계를 거치지 않은 프로젝트는 동시에 《설계임무서》와 《설계기획표》를 배포합니다); 심사 완료 또는 발주자가 확인한 건축 기본설계도를 배포합니다; 각 전문의 기본설계 심사 의견을 배포합니다; 발주자와 정부 부처의 요구를 연구·토론하고 통일된 설계 요구와 조치를 수립합니다; 각 전문은 조사·수집할 자료를 검토하고 가능한 빨리 조사를 조직합니다; 한계 설계는 심사 완료(또는 발주자 동의)된 기본설계 개산을 제출합니다; 본 공사의 《전문배합진도표》를 작성하여 서명 후 효력이 발생합니다",
        "fr": "Organisé par le chef de projet, présidé par le directeur de conception, avec la participation des responsables de toutes les disciplines (au début de la phase de conception des documents d'exécution) : présentation des informations générales du projet (pour les projets n'ayant pas passé par le projet d'esquisse de notre institut, distribution simultanée du Cahier des missions de conception et de la Fiche de planification de conception) ; distribution des plans d'architecture du projet d'esquisse approuvés ou confirmés par le maître d'ouvrage ; distribution des avis d'approbation du projet d'esquisse de toutes les disciplines ; étude et discussion des exigences du maître d'ouvrage et des autorités publiques et formulation d'exigences et de mesures de conception unifiées ; chaque discipline étudie les documents à enquêter et à collecter et organise l'enquête dès que possible ; le design to cost propose l'estimation budgétaire du projet d'esquisse approuvée (ou acceptée par le maître d'ouvrage) ; établissement du Calendrier de coordination professionnelle du présent projet, signé et entré en vigueur",
        "de": "Organisiert durch den Projektleiter, geleitet durch den Entwurfsleiter, mit Teilnahme der Verantwortlichen aller Fachdisziplinen (am Beginn der Ausführungsplanungsphase): Vorstellung der grundlegenden Projektdaten (für Projekte, die nicht durch die Vorplanung unseres Instituts gelaufen sind, gleichzeitig Ausgabe des Entwurfsauftragsbuchs und des Entwurfsplanungsblatts); Ausgabe der genehmigten oder vom Bauherrn bestätigten Architektur-Vorplanungszeichnungen; Ausgabe der Vorplanungsgenehmigungsvermerke aller Disziplinen; Untersuchung und Diskussion der Anforderungen des Bauherrn und der Behörden und Festlegung einheitlicher Entwurfsanforderungen und -maßnahmen; jede Disziplin untersucht die zu erkundenden und zu sammelnden Unterlagen und organisiert die Erkundung so bald wie möglich; Kostenplanung legt die genehmigte (oder vom Bauherrn akzeptierte) Vorplanungskostenschätzung vor; Erstellung des Fachkoordinationszeitplans für dieses Bauvorhaben, Unterzeichnung und Inkraftsetzung",
        "es": "Organizado por el jefe de proyecto, presidido por el director de diseño, con participación de los responsables de todas las especialidades (al inicio de la fase de diseño de documentos de obra): presentación de la información básica del proyecto (para proyectos que no hayan pasado por el diseño preliminar de nuestro instituto, entrega simultánea del Encargo de diseño y de la Hoja de planificación del diseño); entrega de los planos de arquitectura del diseño preliminar aprobados o confirmados por el cliente; entrega de las observaciones de aprobación del diseño preliminar de todas las especialidades; estudio y discusión de los requisitos del cliente y de las autoridades públicas y formulación de requisitos y medidas de diseño unificados; cada especialidad estudia los documentos a investigar y recopilar y organiza la investigación cuanto antes; el diseño al coste propone el presupuesto del diseño preliminar aprobado (o aceptado por el cliente); elaboración del Cronograma de coordinación profesional de esta obra, firmado y en vigor",
        "pt": "Organizado pelo gerente do projeto, presidido pelo diretor de projeto, com a participação dos responsáveis de todas as especialidades (no início da fase de projeto de execução): apresentação das informações básicas do projeto (para projetos que não passaram pelo projeto preliminar do nosso instituto, entrega simultânea do Termo de Tarefa de Projeto e da Ficha de Planejamento de Projeto); entrega dos desenhos de arquitetura do projeto preliminar aprovados ou confirmados pelo cliente; entrega dos pareceres de aprovação do projeto preliminar de todas as especialidades; estudo e discussão dos requisitos do cliente e dos órgãos públicos e formulação de requisitos e medidas unificados de projeto; cada especialidade estuda os documentos a pesquisar e coletar e organiza a pesquisa o mais rápido possível; o design to cost propõe a estimativa orçamentária do projeto preliminar aprovada (ou aceita pelo cliente); elaboração do Cronograma de Coordenação Profissional desta obra, assinado e em vigor",
        "ru": "Организует руководитель проекта, ведёт главный архитектор проекта, участвуют ответственные по всем специальностям (на дату начала стадии «рабочая документация»): представление основных сведений о проекте (для проектов, не проходивших стадию «проект» в нашем институте, одновременно выдаются Задание на проектирование и План проектирования); выдача утверждённых или подтверждённых заказчиком архитектурных чертежей стадии «проект»; выдача замечаний к утверждению стадии «проект» по всем специальностям; изучение и обсуждение требований заказчика и государственных органов и разработка единых требований и мероприятий по проектированию; каждая специальность изучает подлежащие изысканию и сбору материалы и организует изыскания в кратчайшие сроки; проектирование в пределах сметы представляет утверждённую (или согласованную заказчиком) смету стадии «проект»; составление Графика профессиональной координации по данному объекту с подписанием и вступлением в силу",
        "ar": "ينظمه مدير المشروع، ويرأسه مدير التصميم، ويشارك فيه رؤساء جميع التخصصات (في يوم بدء مرحلة تصميم مستندات التنفيذ): تقديم المعلومات الأساسية عن المشروع (بالنسبة للمشاريع التي لم تمر بالتصميم الأولي لمعهدنا، يتم توزيع كتاب مهام التصميم ونموذج تخطيط التصميم في نفس الوقت)؛ وتوزيع مخططات التصميم الأولي المعمارية المعتمدة أو المؤكدة من العميل؛ وتوزيع آراء اعتماد التصميم الأولي لجميع التخصصات؛ ودراسة ومناقشة متطلبات العميل والجهات الحكومية ووضع متطلبات وإجراءات تصميم موحدة؛ ويدرس كل تخصص المواد التي يجب البحث عنها وجمعها وينظم البحث في أقرب وقت؛ ويقدم التصميم وفق التكلفة التقدير المالي للتصميم الأولي المعتمد (أو المقبول من العميل)؛ وإعداد جدول التنسيق المهني لهذا المشروع وتوقيعه ودخوله حيز التنفيذ",
        "it": "Organizzato dal project manager, presieduto dal direttore della progettazione, con la partecipazione dei responsabili di tutte le discipline (all'inizio della fase di progettazione esecutiva): presentazione delle informazioni generali del progetto (per i progetti che non sono passati dalla progettazione preliminare del nostro istituto, distribuzione contestuale del Libretto dei compiti di progettazione e del Foglio di pianificazione della progettazione); distribuzione degli elaborati di progettazione preliminare architettonica approvati o confermati dal committente; distribuzione dei pareri di approvazione della progettazione preliminare di tutte le discipline; studio e discussione dei requisiti del committente e delle autorità pubbliche e definizione di requisiti e misure progettuali unificati; ogni disciplina studia i documenti da indagare e raccogliere e organizza l'indagine al più presto; il design to cost propone il preventivo della progettazione preliminare approvato (o accettato dal committente); predisposizione del Cronoprogramma di coordinamento professionale di questo progetto, firmato ed entrato in vigore",
        "vi": "Do quản lý dự án tổ chức, giám đốc thiết kế chủ trì, các trưởng bộ môn tham dự (ngày bắt đầu giai đoạn thiết kế bản vẽ thi công): giới thiệu thông tin cơ bản của dự án (đối với dự án chưa qua thiết kế cơ sở của viện, đồng thời phát hành Sổ nhiệm vụ thiết kế và Phiếu kế hoạch thiết kế); phát hành bản vẽ thiết kế cơ sở kiến trúc đã được phê duyệt hoặc chủ đầu tư xác nhận; phát hành ý kiến phê duyệt thiết kế cơ sở của các chuyên ngành; nghiên cứu thảo luận yêu cầu của chủ đầu tư và cơ quan nhà nước, đề ra yêu cầu và biện pháp thiết kế thống nhất; mỗi chuyên ngành nghiên cứu tài liệu cần khảo sát thu thập và tổ chức khảo sát sớm nhất có thể; thiết kế theo giới hạn chi phí đề xuất dự toán thiết kế cơ sở đã phê duyệt (hoặc chủ đầu tư đồng ý); lập Bảng tiến độ phối hợp chuyên môn của công trình này và ký để có hiệu lực",
        "th": "จัดโดยผู้จัดการโครงการ โดยผู้อำนวยการฝ่ายออกแบบเป็นประธาน และหัวหน้าทุกสาขาวิชาเข้าร่วม (วันเริ่มต้นระยะการออกแบบแบบก่อสร้าง): แนะนำข้อมูลพื้นฐานของโครงการ (สำหรับโครงการที่ไม่ผ่านการออกแบบขั้นต้นของสถาบันเรา ให้แจกหนังสือภารกิจออกแบบและแบบฟอร์มวางแผนการออกแบบพร้อมกันด้วย) แจกแบบสถาปัตยกรรมขั้นต้นที่ผ่านการอนุมัติหรือยืนยันโดยเจ้าของโครงการ แจกความเห็นการอนุมัติการออกแบบขั้นต้นของทุกสาขา ศึกษาและอภิปรายข้อกำหนดของเจ้าของโครงการและหน่วยงานภาครัฐ กำหนดข้อกำหนดและมาตรการออกแบบที่สอดคล้องกัน แต่ละสาขาศึกษาเอกสารที่ต้องสำรวจและรวบรวมและจัดสำรวจโดยเร็วที่สุด การออกแบบตามต้นทุนเสนอประมาณการงบประมาณการออกแบบขั้นต้นที่ผ่านการอนุมัติ (หรือที่เจ้าของโครงการยอมรับ) จัดทำตารางการประสานงานวิชาชีพของโครงการนี้และลงนามให้มีผลบังคับใช้",
    },
    "建筑专业提出，作为各专业施工图设计准备图。内容：全部初步设计建筑图纸（建筑各层平面图、立面图、剖面图等），平面图立面图剖面图应符合规定2.3.2条要求并满足建设部对初步设计深度的规定；其他需要提出的资料。接受专业：总图、结构、水、暖、电、智能化": {
        "en": "Proposed by the architectural discipline as preparation drawings for the construction document design of all disciplines. Contents: all architectural drawings of the preliminary design (architectural plans of each floor, elevations, sections, etc.); the plans, elevations and sections shall comply with the requirements of Article 2.3.2 and satisfy the provisions of the Ministry of Construction on the depth of preliminary design; other materials to be provided. Receiving disciplines: site planning, structural, plumbing, HVAC, electrical, intelligent systems",
        "ja": "建築専門が提出し、各専門の施工図設計準備図に供する。内容：基本設計の全建築図面（建築各階平面図、立面図、断面図等）。平面図・立面図・断面図は規定2.3.2条の要求を満たし、建設部の基本設計深度の規定を満足する；その他提出すべき資料。受取専門：総図、構造、水、暖、電、智能化",
        "ko": "건축 전문이 제출하며 각 전문의 시공도 설계 준비도면으로 사용됩니다. 내용: 기본설계의 전체 건축 도면(건축 각 층 평면도, 입면도, 단면도 등); 평면도·입면도·단면도는 규정 2.3.2조 요구에 부합하고 건설부의 기본설계 깊이 규정을 충족해야 합니다; 기타 제출해야 할 자료. 수신 전문: 총도, 구조, 급배수, 난방, 전기, 지능화",
        "fr": "Présenté par la discipline architecture, servant de plans préparatoires au projet d'exécution de toutes les disciplines. Contenu : tous les plans d'architecture du projet d'esquisse (plans d'architecture de chaque niveau, élévations, coupes, etc.) ; les plans, élévations et coupes doivent être conformes aux exigences de l'article 2.3.2 et satisfaire aux dispositions du Ministère de la Construction sur la profondeur du projet d'esquisse ; autres documents à fournir. Disciplines destinataires : aménagement du site, structure, plomberie, CVC, électricité, systèmes intelligents",
        "de": "Vorgelegt von der Architekturdisziplin als Vorbereitungszeichnungen für die Ausführungsplanung aller Fachdisziplinen. Inhalt: sämtliche Architekturzeichnungen der Vorplanung (Architekturpläne aller Geschosse, Ansichten, Schnitte usw.); die Grundrisse, Ansichten und Schnitte müssen den Anforderungen von Abschnitt 2.3.2 entsprechen und die Vorgaben des Bauministeriums zur Tiefe der Vorplanung erfüllen; sonstige vorzulegende Unterlagen. Empfangende Disziplinen: Standortplanung, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation",
        "es": "Presentado por la especialidad de arquitectura, como planos preparatorios para el diseño de documentos de obra de todas las especialidades. Contenido: todos los planos de arquitectura del diseño preliminar (plantas de arquitectura de cada nivel, alzados, secciones, etc.); las plantas, alzados y secciones deben cumplir los requisitos del artículo 2.3.2 y satisfacer las disposiciones del Ministerio de Construcción sobre la profundidad del diseño preliminar; demás documentación a aportar. Especialidades receptoras: urbanización, estructuras, fontanería, climatización, electricidad, sistemas inteligentes",
        "pt": "Apresentado pela especialidade de arquitetura, como desenhos preparatórios para o projeto de execução de todas as especialidades. Conteúdo: todos os desenhos de arquitetura do projeto preliminar (plantas de arquitetura de cada pavimento, fachadas, cortes, etc.); as plantas, fachadas e cortes devem atender aos requisitos do artigo 2.3.2 e satisfazer as disposições do Ministério da Construção sobre a profundidade do projeto preliminar; demais documentos a fornecer. Especialidades receptoras: urbanismo, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes",
        "ru": "Представляется архитектурной специальностью в качестве подготовительных чертежей для разработки рабочей документации всеми специальностями. Содержание: все архитектурные чертежи стадии «проект» (архитектурные планы всех этажей, фасады, разрезы и т.п.); планы, фасады и разрезы должны соответствовать требованиям п. 2.3.2 и удовлетворять положениям Министерства строительства о глубине проработки стадии «проект»; прочие материалы, подлежащие представлению. Принимающие специальности: генплан, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы",
        "ar": "يقدمه تخصص العمارة كمخططات تحضيرية لتصميم مستندات التنفيذ لجميع التخصصات. المحتوى: جميع مخططات التصميم الأولي المعمارية (المخططات الأفقية المعمارية لكل طابق والواجهات والقطاعات وغيرها)؛ ويجب أن تتوافق المخططات الأفقية والواجهات والقطاعات مع متطلبات البند 2.3.2 وتفي بأحكام وزارة الإنشاءات بشأن عمق التصميم الأولي؛ وغيرها من المواد التي يلزم تقديمها. التخصصات المستلمة: تخطيط الموقع، الإنشاءات، السباكة، التدفئة والتهوية، الكهرباء، الأنظمة الذكية",
        "it": "Presentato dalla disciplina architettonica come elaborati preparatori per la progettazione esecutiva di tutte le discipline. Contenuto: tutti gli elaborati architettonici della progettazione preliminare (piante architettoniche di ogni piano, prospetti, sezioni, ecc.); le piante, i prospetti e le sezioni devono essere conformi ai requisiti dell'articolo 2.3.2 e soddisfare le disposizioni del Ministero delle Costruzioni sulla profondità della progettazione preliminare; altri documenti da fornire. Discipline destinatarie: pianificazione del sito, strutture, impianti idrici, HVAC, elettrici, sistemi intelligenti",
        "vi": "Do chuyên ngành kiến trúc đưa ra, dùng làm bản vẽ chuẩn bị cho thiết kế bản vẽ thi công của các chuyên ngành. Nội dung: toàn bộ bản vẽ kiến trúc thiết kế cơ sở (mặt bằng kiến trúc từng tầng, mặt đứng, mặt cắt, v.v.); mặt bằng, mặt đứng, mặt cắt phải tuân thủ yêu cầu điều 2.3.2 và đáp ứng quy định của Bộ Xây dựng về độ sâu thiết kế cơ sở; các tài liệu khác cần nêu. Các chuyên ngành nhận: tổng mặt bằng, kết cấu, cấp thoát nước, điều hòa không khí, điện, hệ thống thông minh",
        "th": "เสนอโดยสาขาสถาปัตยกรรม ใช้เป็นแบบเตรียมการออกแบบแบบก่อสร้างของทุกสาขา เนื้อหา: แบบสถาปัตยกรรมทั้งหมดของการออกแบบขั้นต้น (แบบพื้นสถาปัตยกรรมแต่ละชั้น แบบด้าน แบบตัด ฯลฯ) แบบพื้น แบบด้าน แบบตัดต้องเป็นไปตามข้อกำหนดข้อ 2.3.2 และเป็นไปตามระเบียบของกระทรวงการก่อสร้างว่าด้วยความลึกของการออกแบบขั้นต้น เอกสารอื่นที่ต้องเสนอ สาขาที่รับ: การวางผังบริเวณ โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ",
    },
    "总图专业提出，符合规定2.3.3条要求并满足建设部对初步设计深度的规定。接收专业：建筑、结构、水、暖、电、智能化": {
        "en": "Proposed by the site planning discipline, complying with the requirements of Article 2.3.3 and satisfying the provisions of the Ministry of Construction on the depth of preliminary design. Receiving disciplines: architectural, structural, plumbing, HVAC, electrical, intelligent systems",
        "ja": "総図専門が提出し、規定2.3.3条の要求に適合し、建設部の基本設計深度の規定を満足する。受取専門：建築、構造、水、暖、電、智能化",
        "ko": "총도 전문이 제출하며 규정 2.3.3조 요구에 부합하고 건설부의 기본설계 깊이 규정을 충족합니다. 수신 전문: 건축, 구조, 급배수, 난방, 전기, 지능화",
        "fr": "Présenté par la discipline d'aménagement du site, conforme aux exigences de l'article 2.3.3 et satisfaisant aux dispositions du Ministère de la Construction sur la profondeur du projet d'esquisse. Disciplines destinataires : architecture, structure, plomberie, CVC, électricité, systèmes intelligents",
        "de": "Vorgelegt von der Disziplin Standortplanung, in Übereinstimmung mit den Anforderungen von Abschnitt 2.3.3 und Erfüllung der Vorgaben des Bauministeriums zur Tiefe der Vorplanung. Empfangende Disziplinen: Architektur, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation",
        "es": "Presentado por la especialidad de urbanización, conforme a los requisitos del artículo 2.3.3 y satisfaciendo las disposiciones del Ministerio de Construcción sobre la profundidad del diseño preliminar. Especialidades receptoras: arquitectura, estructuras, fontanería, climatización, electricidad, sistemas inteligentes",
        "pt": "Apresentado pela especialidade de urbanismo, em conformidade com os requisitos do artigo 2.3.3 e satisfazendo as disposições do Ministério da Construção sobre a profundidade do projeto preliminar. Especialidades receptoras: arquitetura, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes",
        "ru": "Представляется специальностью генплана, соответствует требованиям п. 2.3.3 и удовлетворяет положениям Министерства строительства о глубине проработки стадии «проект». Принимающие специальности: архитектура, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы",
        "ar": "يقدمه تخصص تخطيط الموقع، بما يتوافق مع متطلبات البند 2.3.3 ويفي بأحكام وزارة الإنشاءات بشأن عمق التصميم الأولي. التخصصات المستلمة: العمارة، الإنشاءات، السباكة، التدفئة والتهوية، الكهرباء، الأنظمة الذكية",
        "it": "Presentato dalla disciplina di pianificazione del sito, conforme ai requisiti dell'articolo 2.3.3 e soddisfacente le disposizioni del Ministero delle Costruzioni sulla profondità della progettazione preliminare. Discipline destinatarie: architettura, strutture, impianti idrici, HVAC, elettrici, sistemi intelligenti",
        "vi": "Do chuyên ngành tổng mặt bằng đưa ra, tuân thủ yêu cầu điều 2.3.3 và đáp ứng quy định của Bộ Xây dựng về độ sâu thiết kế cơ sở. Các chuyên ngành nhận: kiến trúc, kết cấu, cấp thoát nước, điều hòa không khí, điện, hệ thống thông minh",
        "th": "เสนอโดยสาขาการวางผังบริเวณ เป็นไปตามข้อกำหนดข้อ 2.3.3 และเป็นไปตามระเบียบของกระทรวงการก่อสร้างว่าด้วยความลึกของการออกแบบขั้นต้น สาขาที่รับ: สถาปัตยกรรม โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ",
    },
    "结构专业提给建筑专业：考虑结构方案，针对建筑初步设计图研究结构方案的可行性；提出地质勘察要求；提出结构构造措施（伸缩缝、沉降缝、防震缝位置要求及特殊部位构造做法）；提出人工地基处理方法、处理深度；提出结构选型；提出基础形式及基础截面尺寸；对建筑方案修改的其他建议": {
        "en": "Structural discipline to architectural discipline: consider the structural scheme and study the feasibility of the structural scheme against the architectural preliminary design drawings; raise the geotechnical survey requirements; raise structural construction measures (position requirements of expansion joints, settlement joints and seismic joints and construction practices for special parts); raise the treatment methods and treatment depth of artificial ground; raise the structural system selection; raise the foundation forms and foundation section dimensions; other suggestions for modifying the architectural scheme",
        "ja": "構造専門から建築専門へ：構造方案を考慮し、建築基本設計図に対して構造方案の実現可能性を研究する；地質調査の要求を提出する；構造上の措置（伸縮縫、沈下縫、防震縫の位置要求および特殊部位の構造仕様）を提出する；人工地盤の処理方法・処理深度を提出する；構造システムの選定を提出する；基礎形式および基礎断面寸法を提出する；建築方案の修正に関するその他の提案",
        "ko": "구조 전문이 건축 전문에게: 구조 안을 고려하여 건축 기본설계도에 대해 구조 안의 실현 가능성을 연구합니다; 지질 조사 요구를 제출합니다; 구조 조치(신축줄눈, 침하줄눈, 방진줄눈 위치 요구 및 특수 부위 구조 공법)를 제출합니다; 인공지반 처리 방법·처리 깊이를 제출합니다; 구조 시스템 선정을 제출합니다; 기초 형식 및 기초 단면 치수를 제출합니다; 건축 안 수정에 관한 기타 제안",
        "fr": "Structure à architecture : considérer le schéma de structure et étudier la faisabilité du schéma de structure au regard des plans d'architecture du projet d'esquisse ; formuler les exigences de reconnaissance géotechnique ; formuler les mesures constructives structurelles (exigences de position des joints de dilatation, de tassement et parasismiques et pratiques constructives des parties spéciales) ; formuler les méthodes et la profondeur de traitement du sol artificiel ; formuler le choix du système structurel ; formuler les formes de fondations et les dimensions de leurs sections ; autres suggestions de modification du schéma d'architecture",
        "de": "Tragwerk an Architektur: unter Berücksichtigung des Tragwerkskonzepts die Machbarkeit des Tragwerkskonzepts anhand der Architektur-Vorplanungszeichnungen untersuchen; Anforderungen an die Baugrunduntersuchung stellen; konstruktive Maßnahmen vorschlagen (Lageanforderungen für Dehnungs-, Setzungs- und Erdbebenfugen sowie Konstruktionsarten besonderer Teile); Behandlungsmethoden und -tiefe des künstlichen Baugrunds vorschlagen; die Wahl des Tragwerkssystems vorschlagen; Gründungsformen und Querschnittsabmessungen der Gründungen vorschlagen; sonstige Vorschläge zur Änderung des Architekturkonzepts",
        "es": "Estructuras a arquitectura: considerar el esquema estructural y estudiar la viabilidad del esquema estructural respecto a los planos de arquitectura del diseño preliminar; plantear los requisitos de reconocimiento geotécnico; plantear las medidas constructivas estructurales (requisitos de posición de juntas de dilatación, asiento y antisísmicas y prácticas constructivas de partes especiales); plantear los métodos y la profundidad de tratamiento del terreno artificial; plantear la selección del sistema estructural; plantear las formas de cimentación y las dimensiones de sus secciones; otras sugerencias para modificar el esquema de arquitectura",
        "pt": "Estruturas para arquitetura: considerar o esquema estrutural e estudar a viabilidade do esquema estrutural em relação aos desenhos de arquitetura do projeto preliminar; apresentar os requisitos de investigação geotécnica; apresentar as medidas construtivas estruturais (requisitos de posição de juntas de dilatação, de recalque e sísmicas e práticas construtivas de partes especiais); apresentar os métodos e a profundidade de tratamento do terreno artificial; apresentar a seleção do sistema estrutural; apresentar as formas de fundação e as dimensões de suas seções; outras sugestões para modificar o esquema arquitetônico",
        "ru": "Конструктивная специальность — архитектурной: с учётом конструктивного решения изучить его осуществимость по архитектурным чертежам стадии «проект»; предъявить требования к инженерно-геологическим изысканиям; представить конструктивные мероприятия (требования к расположению деформационных, осадочных и антисейсмических швов и конструктивные решения особых участков); представить способы и глубину обработки искусственного основания; представить выбор конструктивной системы; представить формы фундаментов и размеры их сечений; прочие предложения по корректировке архитектурных решений",
        "ar": "الإنشاءات إلى العمارة: النظر في التصميم الإنشائي ودراسة جدوى التصميم الإنشائي مقابل مخططات التصميم الأولي المعمارية؛ وطرح متطلبات الاستكشاف الجيوتقني؛ وطرح إجراءات الإنشاء (متطلبات مواقع فواصل التمدد وفواصل الهبوط وفواصل مقاومة الزلازل وأساليب إنشاء الأجزاء الخاصة)؛ وطرح طرق وعمق معالجة التربة الاصطناعية؛ وطرح اختيار النظام الإنشائي؛ وطرح أشكال الأساسات وأبعاد مقاطعها؛ وتقديم اقتراحات أخرى لتعديل التصميم المعماري",
        "it": "Strutture ad architettura: considerare lo schema strutturale e studiarne la fattibilità sulla base degli elaborati architettonici di progettazione preliminare; formulare i requisiti delle indagini geotecniche; formulare le misure costruttive strutturali (requisiti di posizione dei giunti di dilatazione, di assestamento e sismici e pratiche costruttive delle parti speciali); formulare i metodi e la profondità di trattamento del terreno artificiale; formulare la scelta del sistema strutturale; formulare le forme di fondazione e le dimensioni delle sezioni; altri suggerimenti per modificare lo schema architettonico",
        "vi": "Kết cấu gửi kiến trúc: xem xét phương án kết cấu, nghiên cứu tính khả thi của phương án kết cấu đối chiếu bản vẽ thiết kế cơ sở kiến trúc; nêu yêu cầu khảo sát địa chất; nêu biện pháp cấu tạo kết cấu (yêu cầu vị trí khe co giãn, khe lún, khe chống động đất và biện pháp cấu tạo bộ phận đặc biệt); nêu phương pháp và chiều sâu xử lý nền nhân tạo; nêu lựa chọn hệ kết cấu; nêu hình thức móng và kích thước tiết diện móng; các đề xuất khác sửa đổi phương án kiến trúc",
        "th": "โครงสร้างส่งให้สถาปัตยกรรม: พิจารณาแบบโครงสร้างและศึกษาความเป็นไปได้ของแบบโครงสร้างเทียบกับแบบสถาปัตยกรรมขั้นต้น เสนอข้อกำหนดการสำรวจธรณีเทคนิค เสนอมาตรการก่อสร้างเชิงโครงสร้าง (ข้อกำหนดตำแหน่งรอยต่อการขยายตัว รอยต่อทรุดตัว รอยต่อต้านแผ่นดินไหว และวิธีก่อสร้างส่วนพิเศษ) เสนอวิธีการและความลึกของการปรับปรุงฐานรากเทียม เสนอการเลือกระบบโครงสร้าง เสนอรูปแบบฐานรากและขนาดหน้าตัดฐานราก ข้อเสนออื่น ๆ ในการปรับปรุงแบบสถาปัตยกรรม",
    },
    "机电专业提给建筑专业：考虑机电系统设计方案，针对建筑初步设计图研究对机电系统的要求，特别是现有初步设计的机房、管道井位置和面积的可行性，提出修改意见；提出新增加的机房和管道井的大致面积和初步位置；大型设备吊装孔位置和尺寸要求及运输通道要求；对建筑方案修改的其他建议": {
        "en": "MEP disciplines to architectural discipline: consider the MEP system design scheme and study, against the architectural preliminary design drawings, the requirements for the MEP systems, in particular the feasibility of the positions and areas of the plant rooms and riser shafts in the existing preliminary design, and raise modification opinions; raise the approximate areas and preliminary positions of newly added plant rooms and riser shafts; raise the position and size requirements of lifting openings for large equipment and the requirements for transport routes; other suggestions for modifying the architectural scheme",
        "ja": "機電専門から建築専門へ：機電システムの設計方案を考慮し、建築基本設計図に対して機電システムへの要求、特に現行基本設計の機械室・パイプシャフトの位置と面積の実現可能性を研究し、修正意見を提出する；新設する機械室・パイプシャフトのおおよその面積と初期位置を提出する；大型設備の吊込孔の位置・寸法要求および運搬通路の要求を提出する；建築方案の修正に関するその他の提案",
        "ko": "기전 전문이 건축 전문에게: 기전 시스템 설계 안을 고려하여 건축 기본설계도에 대해 기전 시스템 요구, 특히 기존 기본설계의 기계실·파이프 샤프트 위치와 면적의 실현 가능성을 연구하고 수정 의견을 제출합니다; 신설되는 기계실·파이프 샤프트의 대략적인 면적과 초기 위치를 제출합니다; 대형 설비 인양 구멍의 위치·치수 요구 및 운반 동선 요구를 제출합니다; 건축 안 수정에 관한 기타 제안",
        "fr": "MEP à architecture : considérer le schéma de conception des systèmes CVC-électricité-plomberie et étudier, au regard des plans d'architecture du projet d'esquisse, les exigences relatives aux systèmes MEP, en particulier la faisabilité des positions et des surfaces des locaux techniques et des gaines dans le projet d'esquisse existant, et formuler des avis de modification ; formuler les surfaces approximatives et les positions préliminaires des locaux techniques et gaines nouvellement ajoutés ; formuler les exigences de position et de dimensions des ouvertures de levage des grands équipements et les exigences des voies de transport ; autres suggestions de modification du schéma d'architecture",
        "de": "TGA an Architektur: unter Berücksichtigung des TGA-Entwurfskonzepts anhand der Architektur-Vorplanungszeichnungen die Anforderungen an die TGA-Systeme untersuchen, insbesondere die Machbarkeit der Positionen und Flächen der Technikräume und Schächte in der bestehenden Vorplanung, und Änderungsvorschläge unterbreiten; ungefähre Flächen und vorläufige Positionen neu hinzugefügter Technikräume und Schächte vorschlagen; Positions- und Größenvorgaben für Hebeöffnungen großer Geräte sowie Anforderungen an Transportwege vorschlagen; sonstige Vorschläge zur Änderung des Architekturkonzepts",
        "es": "MEP a arquitectura: considerar el esquema de diseño de los sistemas MEP y estudiar, respecto a los planos de arquitectura del diseño preliminar, los requisitos para los sistemas MEP, en particular la viabilidad de las posiciones y superficies de los locales técnicos y patinillos del diseño preliminar existente, y formular observaciones de modificación; formular las superficies aproximadas y las posiciones preliminares de los locales técnicos y patinillos añadidos; formular los requisitos de posición y dimensiones de las aberturas de izado de los grandes equipos y los requisitos de las vías de transporte; otras sugerencias para modificar el esquema de arquitectura",
        "pt": "MEP para arquitetura: considerar o esquema de projeto dos sistemas MEP e estudar, em relação aos desenhos de arquitetura do projeto preliminar, os requisitos para os sistemas MEP, em particular a viabilidade das posições e áreas das salas técnicas e shaft do projeto preliminar existente, e formular observações de modificação; formular as áreas aproximadas e as posições preliminares das salas técnicas e shaft recém-adicionados; formular os requisitos de posição e dimensões das aberturas de içamento dos grandes equipamentos e os requisitos das vias de transporte; outras sugestões para modificar o esquema arquitetônico",
        "ru": "Инженерные специальности — архитектурной: с учётом проектных решений инженерных систем изучить по архитектурным чертежам стадии «проект» требования к инженерным системам, в особенности осуществимость расположения и площадей технических помещений и шахт в существующем проекте, и представить замечания по корректировке; представить ориентировочные площади и предварительное расположение вновь добавляемых технических помещений и шахт; представить требования к расположению и размерам монтажных проёмов для крупного оборудования и требования к транспортным путям; прочие предложения по корректировке архитектурных решений",
        "ar": "التخصصات الميكانيكية والكهربائية والسباكة إلى العمارة: النظر في مخطط تصميم أنظمة الأعمال الميكانيكية والكهربائية والسباكة ودراسة، مقابل مخططات التصميم الأولي المعمارية، متطلبات أنظمة الأعمال الميكانيكية والكهربائية والسباكة، ولا سيما جدوى مواقع ومساحات الغرف الفنية والأعمدة في التصميم الأولي الحالي، وتقديم آراء التعديل؛ وتقديم المساحات التقريبية والمواقع الأولية للغرف الفنية والأعمدة المضافة حديثاً؛ وتقديم متطلبات موقع وأبعاد فتحات رفع المعدات الكبيرة ومتطلبات طرق النقل؛ وتقديم اقتراحات أخرى لتعديل التصميم المعماري",
        "it": "MEP ad architettura: considerare lo schema progettuale dei sistemi MEP e studiare, sulla base degli elaborati architettonici di progettazione preliminare, i requisiti per i sistemi MEP, in particolare la fattibilità di posizioni e superfici dei locali tecnici e dei vani tecnici nella progettazione preliminare esistente, e formulare pareri di modifica; formulare le superfici approssimative e le posizioni preliminari dei locali tecnici e dei vani tecnici di nuova aggiunta; formulare i requisiti di posizione e dimensioni delle aperture di sollevamento per le grandi apparecchiature e i requisiti dei percorsi di trasporto; altri suggerimenti per modificare lo schema architettonico",
        "vi": "Cơ điện gửi kiến trúc: xem xét phương án thiết kế hệ thống cơ điện, nghiên cứu đối chiếu bản vẽ thiết kế cơ sở kiến trúc các yêu cầu đối với hệ thống cơ điện, đặc biệt là tính khả thi về vị trí, diện tích phòng máy và giếng kỹ thuật trong thiết kế cơ sở hiện có, và nêu ý kiến sửa đổi; nêu diện tích gần đúng và vị trí sơ bộ của phòng máy, giếng kỹ thuật bổ sung mới; nêu yêu cầu vị trí, kích thước lỗ cẩu thiết bị lớn và yêu cầu đường vận chuyển; các đề xuất khác sửa đổi phương án kiến trúc",
        "th": "งานระบบเครื่องกลไฟฟ้าส่งให้สถาปัตยกรรม: พิจารณาแบบการออกแบบระบบงานเครื่องกลไฟฟ้า ศึกษาเทียบกับแบบสถาปัตยกรรมขั้นต้นถึงข้อกำหนดสำหรับระบบงานเครื่องกลไฟฟ้า โดยเฉพาะความเป็นไปได้ของตำแหน่งและพื้นที่ห้องเครื่องและปล่องท่อในการออกแบบขั้นต้นปัจจุบัน และเสนอความเห็นการแก้ไข เสนอพื้นที่โดยประมาณและตำแหน่งเบื้องต้นของห้องเครื่องและปล่องท่อที่เพิ่มใหม่ เสนอข้อกำหนดตำแหน่งและขนาดช่องยกอุปกรณ์ขนาดใหญ่และข้อกำหนดเส้นทางขนส่ง ข้อเสนออื่น ๆ ในการปรับปรุงแบบสถาปัตยกรรม",
    },
    "建筑专业提出，作为各专业施工图设计工作基础用图。内容（在第一阶段基础上）：反映初步设计审批意见和第一阶段配合成果；各层建筑主要平面图、立面图及剖面图（人防平面及口部布置、防火分区、机电能源站房及分散机房位置、主要建筑构配件、主要建筑设备位置、房间名称与设计标高、二次装修区域及吊顶净高、主要管道井位置、屋顶设备与屋面绿化位置等）；文字说明（楼地面墙体屋面及地下室底板构造做法、不同功能区域建筑面积和（或）设计使用人数）；必要的初步详图和放大图（客房、住宅单元放大图）；标明外门、外窗尺寸及玻璃初步性能要求（光学、热工性能）；室外照明设计的初步要求。接受专业：总图、结构、水、暖、电、智能化、经济": {
        "en": "Proposed by the architectural discipline as the base drawings for the construction document design work of all disciplines. Contents (on the basis of Stage 1): reflect the approval opinions of the preliminary design and the coordination results of Stage 1; the main floor plans, elevations and sections of each floor of the building (civil air defense plans and opening layouts, fire compartmentation, positions of mechanical and electrical energy station rooms and decentralized plant rooms, main building components, positions of main building equipment, room names and design elevations, secondary renovation areas and ceiling clear heights, positions of main pipe shafts, roof equipment and roof greening positions, etc.); text descriptions (construction practices of floors, walls, roofs and basement bottom slabs, building areas and/or design occupant numbers of different functional zones); necessary preliminary detail drawings and enlarged drawings (enlarged drawings of guest rooms and residential units); indicate the dimensions of external doors and windows and the preliminary performance requirements of glass (optical and thermal performance); preliminary requirements for outdoor lighting design. Receiving disciplines: site planning, structural, plumbing, HVAC, electrical, intelligent systems, economics",
        "ja": "建築専門が提出し、各専門の施工図設計作業の基礎図とする。内容（第一段階を基礎とする）：基本設計の審査意見と第一段階の配合成果を反映する；各階の建築主要平面図・立面図・断面図（人防平面と口部配置、防火区画、電気機械エネルギー源ステーションと分散機械室の位置、主要建築部品・構件、主要建築設備の位置、室名と設計標高、二次装修エリアと吊り天井高、主要パイプシャフト位置、屋上設備と屋上緑化位置等）；文字説明（床・壁・屋根・地下室外壁・底板の構造仕様、異なる機能区域の建築面積および（または）設計使用人数）；必要な初期詳細図と拡大図（客室・住宅ユニット拡大図）；外扉・外窓の寸法とガラスの初期性能要求（光学・熱工性能）を明示する；屋外照明設計の初期要求。受取専門：総図、構造、水、暖、電、智能化、経済",
        "ko": "건축 전문이 제출하며 각 전문의 시공도 설계 작업의 기초 도면으로 사용합니다. 내용(제1단계를 기반으로): 기본설계 심사 의견과 제1단계 배합 성과를 반영합니다; 각 층 건축 주요 평면도·입면도·단면도(인방 평면과 출입구 배치, 방화구획, 전기·기계 에너지 스테이션과 분산 기계실 위치, 주요 건축 부품·부재, 주요 건축 설비 위치, 실명과 설계 표고, 2차 실내장식 구역과 천장고, 주요 파이프 샤프트 위치, 옥상 설비와 옥상 녹화 위치 등); 문자 설명(바닥·벽·지붕·지하실 밑판의 구조 시공법, 기능 구역별 건축 면적 및(또는) 설계 사용 인원수); 필요한 초기 상세도와 확대도(객실·주거 유닛 확대도); 외문·외창의 치수와 유리의 초기 성능 요구(광학·열공 성능)를 명시합니다; 옥외 조명 설계의 초기 요구. 수신 전문: 총도, 구조, 급배수, 난방, 전기, 지능화, 경제",
        "fr": "Présenté par la discipline architecture comme fonds de plan pour les travaux de conception des documents d'exécution de toutes les disciplines. Contenu (sur la base de la phase 1) : refléter les avis d'approbation du projet d'esquisse et les résultats de coordination de la phase 1 ; les plans principaux, élévations et coupes de chaque niveau (plans de défense civile et disposition des ouvertures, compartimentage coupe-feu, positions des locaux techniques énergétiques et des locaux techniques décentralisés, principaux composants du bâtiment, positions des principaux équipements du bâtiment, noms des pièces et cotes de niveau, zones de rénovation secondaire et hauteurs sous plafond, positions des gaines techniques principales, positions des équipements de toiture et de la végétalisation de toiture, etc.) ; les notes écrites (modes constructifs des sols, murs, toitures et de la dalle de fond du sous-sol, surfaces de plancher et/ou nombres d'occupants de conception des différentes zones fonctionnelles) ; les avant-projets de détails et les agrandissements nécessaires (agrandissements des chambres et des unités résidentielles) ; indiquer les dimensions des portes et fenêtres extérieures et les exigences préliminaires de performance du verre (performances optiques et thermiques) ; les exigences préliminaires pour la conception de l'éclairage extérieur. Disciplines destinataires : aménagement du site, structure, plomberie, CVC, électricité, systèmes intelligents, économie",
        "de": "Vorgelegt von der Architekturdisziplin als Grundlage für die Ausführungsplanung aller Fachdisziplinen. Inhalt (auf der Grundlage von Stufe 1): die Genehmigungsauflagen der Vorplanung und die Koordinationsergebnisse von Stufe 1 widerspiegeln; die Hauptgrundrisse, Ansichten und Schnitte jedes Geschosses (Schutzraumpäne und Öffnungsanordnungen, Brandabschnitte, Positionen der Energiezentralen und dezentralen Technikräume, wichtigste Gebäudebauteile, Positionen der wichtigsten Gebäudetechnik, Raumnamen und Planungshöhen, Bereiche der Sekundärausbaus und lichte Deckenhöhen, Positionen der Hauptschächte, Dachgerätepositionen und Dachbegrünungsflächen usw.); Erläuterungen (Konstruktionsarten von Fußböden, Wänden, Dächern und der Kellerbodenplatte, Geschossflächen und/oder geplante Nutzerzahlen verschiedener Funktionsbereiche); notwendige vorläufige Detail- und Vergrößerungszeichnungen (Vergrößerungen von Gästezimmern und Wohneinheiten); Abmessungen der Außentüren und -fenster und vorläufige Leistungsanforderungen der Verglasung (optische und thermische Leistung) angeben; vorläufige Anforderungen an die Außenbeleuchtung. Empfangende Disziplinen: Standortplanung, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de arquitectura como base para el diseño de los documentos de obra de todas las especialidades. Contenido (sobre la base de la fase 1): reflejar los dictámenes de aprobación del diseño preliminar y los resultados de coordinación de la fase 1; los planos principales de planta, alzados y secciones de cada nivel (planos de defensa civil y disposición de aberturas, compartimentación contra incendios, posiciones de las estaciones de energía de instalaciones y de los locales técnicos descentralizados, principales componentes del edificio, posiciones de los principales equipos del edificio, nombres de los locales y cotas de nivel, zonas de reforma secundaria y alturas libres bajo techo, posiciones de los patinillos técnicos principales, posiciones de los equipos de cubierta y de las zonas verdes en cubierta, etc.); notas escritas (modos constructivos de suelos, muros, cubiertas y de la solera del sótano, superficies construidas y/o números de ocupantes de diseño de las distintas zonas funcionales); los anteproyectos de detalles y alzados ampliados necesarios (ampliados de habitaciones y unidades residenciales); indicar las dimensiones de puertas y ventanas exteriores y los requisitos preliminares de prestaciones del vidrio (prestaciones ópticas y térmicas); los requisitos preliminares para el diseño de la iluminación exterior. Especialidades receptoras: urbanización, estructuras, fontanería, climatización, electricidad, sistemas inteligentes, economía",
        "pt": "Apresentado pela especialidade de arquitetura como base para o projeto de execução de todas as especialidades. Conteúdo (com base na fase 1): refletir os pareceres de aprovação do projeto preliminar e os resultados de coordenação da fase 1; as plantas principais, elevações e cortes de cada pavimento (plantas de defesa civil e disposição das aberturas, compartimentação contra incêndio, posições das centrais de energia das instalações e das salas técnicas descentralizadas, principais componentes do edifício, posições dos principais equipamentos do edifício, nomes dos ambientes e cotas de nível, áreas de reforma secundária e pés-direitos, posições dos shaft técnicos principais, posições dos equipamentos de cobertura e das áreas verdes de cobertura, etc.); notas escritas (modos construtivos de pisos, paredes, coberturas e da laje de fundo do subsolo, áreas construídas e/ou números de ocupantes de projeto das diferentes zonas funcionais); os anteprojetos de detalhes e ampliações necessários (ampliações de quartos de hóspedes e unidades residenciais); indicar as dimensões de portas e janelas externas e os requisitos preliminares de desempenho do vidro (desempenho óptico e térmico); os requisitos preliminares para o projeto de iluminação externa. Especialidades receptoras: urbanismo, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes, economia",
        "ru": "Представляется архитектурной специальностью в качестве основы для разработки рабочей документации всеми специальностями. Содержание (на основе 1-го этапа): отразить решения по утверждению стадии «проект» и результаты координации 1-го этапа; основные планы этажей, фасады и разрезы (планы гражданской обороны с расположением входов, противопожарные отсеки, расположение энергетических станций инженерных систем и децентрализованных технических помещений, основные строительные конструкции и детали, расположение основного инженерного оборудования, наименования помещений и проектные отметки, зоны вторичной отделки и высоты подвесных потолков, расположение основных шахт, расположение оборудования на кровле и озеленение кровли и т.п.); пояснительные записки (конструктивные решения полов, стен, кровли и плиты пола подвала, площади и (или) проектное количество людей для различных функциональных зон); необходимые предварительные детали и укрупнённые чертежи (укрупнённые планы гостиничных номеров и квартир); указать размеры наружных дверей и окон и предварительные требования к характеристикам остекления (оптические и теплотехнические); предварительные требования к проектированию наружного освещения. Принимающие специальности: генплан, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы, экономика",
        "ar": "يقدمه تخصص العمارة كأساس لأعمال تصميم مستندات التنفيذ لجميع التخصصات. المحتوى (استناداً إلى المرحلة الأولى): عكس آراء اعتماد التصميم الأولي ونتائج التنسيق للمرحلة الأولى؛ والمخططات الرئيسية والواجهات والقطاعات لكل طابق (مخططات الدفاع المدني وترتيب الفتحات، وتقسيمات الحماية من الحرائق، ومواقع محطات الطاقة للأنظمة الكهروميكانيكية والغرف الفنية الموزعة، والمكونات المعمارية الرئيسية، ومواقع المعدات الرئيسية للمبنى، وأسماء الغرف ومناسيب التصميم، ومناطق التشطيب الثانوي وارتفاعات الأسقف المعلقة، ومواقع الأعمدة الفنية الرئيسية، ومواقع معدات الأسطح والمساحات الخضراء على الأسطح وغيرها)؛ والملاحظات النصية (طرق إنشاء الأرضيات والجدران والأسقف واللوح السفلي للطابق السفلي، ومساحات البناء و/أو أعداد المستخدمين التصميمية للمناطق الوظيفية المختلفة)؛ والتفاصيل الأولية اللازمة والتكبيرات (تكبيرات غرف الضيوف والوحدات السكنية)؛ وبيان أبعاد الأبواب والنوافذ الخارجية ومتطلبات الأداء الأولية للزجاج (الأداء البصري والحراري)؛ والمتطلبات الأولية لتصميم الإضاءة الخارجية. التخصصات المستلمة: تخطيط الموقع، الإنشاءات، السباكة، التدفئة والتهوية، الكهرباء، الأنظمة الذكية، الاقتصاد",
        "it": "Presentato dalla disciplina architettonica come base per il progetto esecutivo di tutte le discipline. Contenuto (sulla base della fase 1): riflettere i pareri di approvazione della progettazione preliminare e i risultati di coordinamento della fase 1; le piante principali, prospetti e sezioni di ogni piano (piante di difesa civile e disposizione delle aperture, compartimentazione antincendio, posizioni delle centrali energetiche degli impianti e dei locali tecnici decentralizzati, principali componenti dell'edificio, posizioni delle principali attrezzature dell'edificio, nomi dei locali e quote di progetto, aree di ristrutturazione secondaria e altezze dei controsoffitti, posizioni dei vani tecnici principali, posizioni delle attrezzature di copertura e delle aree verdi di copertura, ecc.); note scritte (modalità costruttive di pavimenti, pareti, coperture e della soletta di fondo del seminterrato, superfici costruite e/o numeri di occupanti di progetto delle diverse zone funzionali); i dettagli preliminari necessari e gli ingrandimenti (ingrandimenti di camere per ospiti e unità residenziali); indicare le dimensioni di porte e finestre esterne e i requisiti prestazionali preliminari del vetro (prestazioni ottiche e termiche); i requisiti preliminari per la progettazione dell'illuminazione esterna. Discipline destinatarie: pianificazione del sito, strutture, idrici, HVAC, elettrici, sistemi intelligenti, economia",
        "vi": "Chuyên ngành kiến trúc đưa ra, dùng làm bản vẽ cơ sở cho công tác thiết kế bản vẽ thi công của các chuyên ngành. Nội dung (trên cơ sở giai đoạn 1): phản ánh ý kiến phê duyệt thiết kế cơ sở và kết quả phối hợp giai đoạn 1; mặt bằng chính, mặt đứng, mặt cắt từng tầng (mặt bằng phòng hộ dân sự và bố trí cửa mở, phân vùng chống cháy, vị trí trạm năng lượng cơ điện và phòng máy phân tán, cấu kiện kiến trúc chính, vị trí thiết bị chính của tòa nhà, tên phòng và cao độ thiết kế, khu vực trang trí thứ cấp và chiều cao trần treo, vị trí giếng kỹ thuật chính, vị trí thiết bị mái và cây xanh mái, v.v.); thuyết minh (cấu tạo nền, tường, mái và bản đáy hầm, diện tích sàn và/hoặc số người sử dụng thiết kế của các khu chức năng khác nhau); các bản vẽ chi tiết và phóng to sơ bộ cần thiết (phóng to phòng khách, đơn nguyên nhà ở); ghi rõ kích thước cửa đi, cửa sổ ngoài và yêu cầu tính năng sơ bộ của kính (quang học, nhiệt); yêu cầu sơ bộ thiết kế chiếu sáng ngoài trời. Các chuyên ngành nhận: tổng mặt bằng, kết cấu, cấp thoát nước, điều hòa, điện, hệ thống thông minh, kinh tế",
        "th": "สาขาสถาปัตยกรรมนำเสนอ ใช้เป็นแบบพื้นฐานสำหรับงานออกแบบแบบก่อสร้างของทุกสาขา เนื้อหา (บนพื้นฐานระยะที่ 1): สะท้อนความเห็นการอนุมัติการออกแบบขั้นต้นและผลการประสานงานระยะที่ 1 แบบแปลนหลัก แบบ立面 และแบบตัดของแต่ละชั้น (แบบป้องกันภัยทางอากาศและการจัดวางช่องเปิด การแบ่งเขตอัคคีภัย ตำแหน่งสถานีพลังงานระบบกลไกไฟฟ้าและห้องเครื่องกระจาย ตำแหน่งส่วนประกอบหลักของอาคาร ตำแหน่งอุปกรณ์หลักของอาคาร ชื่อห้องและระดับออกแบบ พื้นที่ตกแต่งขั้นที่สองและความสูงเพดานแบบแขวน ตำแหน่งปล่องเทคนิคหลัก ตำแหน่งอุปกรณ์บนหลังคาและพื้นที่เขียวบนหลังคา ฯลฯ) คำอธิบายข้อความ (วิธีก่อสร้างพื้น ผนัง หลังคา และแผ่นพื้นใต้ดิน พื้นที่อาคารและ/หรือจำนวนผู้ใช้ตามการออกแบบของเขตหน้าที่ต่าง ๆ) แบบรายละเอียดและภาพขยายเบื้องต้นที่จำเป็น (ภาพขยายห้องพักและชุดที่อยู่อาศัย) ระบุขนาดประตูหน้าต่างภายนอกและข้อกำหนดสมรรถนะเบื้องต้นของกระจก (เชิงแสง เชิงความร้อน) ข้อกำหนดเบื้องต้นสำหรับการออกแบบแสงสว่างภายนอก สาขาที่รับ: การวางผังบริเวณ โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ เศรษฐศาสตร์",
    },
    "水提暖：特殊房间温湿度参数和通风要求（如水箱间温度要求、气体消防房间名称等）；暖提水：不采暖房间名称；暖提电：需要防雷接地和防静电的设备名称（油气锅炉房、烟囱、燃气放散管、可燃气体管道等）；电提水：特殊房间消防要求；电提暖：特殊房间温湿度参数和通风要求；智能化提暖：各弱电机房、安防控制室温湿度参数和通风要求；智能化提电：各弱电机房、安防控制室用电量": {
        "en": "Plumbing to HVAC: temperature and humidity parameters and ventilation requirements of special rooms (such as water tank room temperature requirements, names of gas fire suppression rooms, etc.); HVAC to plumbing: names of unheated rooms; HVAC to electrical: names of equipment requiring lightning protection, earthing and anti-static measures (oil and gas boiler rooms, chimneys, gas vent pipes, combustible gas pipelines, etc.); electrical to plumbing: fire protection requirements for special rooms; electrical to HVAC: temperature and humidity parameters and ventilation requirements of special rooms; intelligent systems to HVAC: temperature and humidity parameters and ventilation requirements of each weak-current room and security control room; intelligent systems to electrical: power consumption of each weak-current room and security control room",
        "ja": "水提暖：特殊な部屋の温湿度パラメータと換気要求（例：水箱室の温度要求、ガス消火室の名称等）；暖提水：非暖房室の名称；暖提電：防雷接地・防静電を要する設備名（油ガスボイラー室・煙突・ガス放散管・可燃性ガス配管等）；電提水：特殊な部屋の消防要求；電提暖：特殊な部屋の温湿度パラメータと換気要求；智能化提暖：各弱電室・安防控制室の温湿度パラメータと換気要求；智能化提電：各弱電室・安防控制室の電力消費量",
        "ko": "급배수가 난방에: 특수실의 온습도 파라미터와 환기 요구(예: 물탱크실 온도 요구, 가스 소방실 명칭 등); 난방이 급배수에: 비난방실 명칭; 난방이 전기에: 방뢰 접지·정전기 방지가 필요한 설비명(유가스 보일러실·굴뚝·가스 방산관·가연성 가스 배관 등); 전기가 급배수에: 특수실 소방 요구; 전기가 난방에: 특수실 온습도 파라미터와 환기 요구; 지능화가 난방에: 각 약전실·안보 통제실의 온습도 파라미터와 환기 요구; 지능화가 전기에: 각 약전실·안보 통제실의 전력 소비량",
        "fr": "Plomberie à CVC : paramètres de température et d'humidité et exigences de ventilation des pièces spéciales (telles que les exigences de température du local de réservoir d'eau, les noms des pièces de protection incendie par gaz, etc.) ; CVC à plomberie : noms des pièces non chauffées ; CVC à électricité : noms des équipements nécessitant une mise à la terre de protection contre la foudre et antistatique (chaufferies gazole/gaz, cheminées, conduits de vente de gaz, canalisations de gaz combustible, etc.) ; électricité à plomberie : exigences de protection incendie des pièces spéciales ; électricité à CVC : paramètres de température et d'humidité et exigences de ventilation des pièces spéciales ; systèmes intelligents à CVC : paramètres de température et d'humidité et exigences de ventilation de chaque local courants faibles et de la salle de contrôle de sécurité ; systèmes intelligents à électricité : consommation électrique de chaque local courants faibles et de la salle de contrôle de sécurité",
        "de": "Sanitär an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter und Lüftungsanforderungen besonderer Räume (z. B. Temperaturanforderungen des Behälterraums, Namen von Gaslöschräumen usw.); Heizung-Lüftung an Sanitär: Namen unbeheizter Räume; Heizung-Lüftung an Elektro: Namen der Anlagen mit Anforderungen an Blitzschutzerdung und Antistatik (Öl- und Gasheizräume, Schornsteine, Gasabblaseleitungen, brennbare Gasleitungen usw.); Elektro an Sanitär: Brandschutzanforderungen für besondere Räume; Elektro an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter und Lüftungsanforderungen besonderer Räume; Gebäudeautomation an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter und Lüftungsanforderungen jedes Schwachstromraums und des Sicherheitskontrollraums; Gebäudeautomation an Elektro: Stromverbrauch jedes Schwachstromraums und des Sicherheitskontrollraums",
        "es": "Fontanería a climatización: parámetros de temperatura y humedad y requisitos de ventilación de los locales especiales (como los requisitos de temperatura del local de depósito de agua, los nombres de los locales de protección contra incendios por gas, etc.); climatización a fontanería: nombres de los locales sin calefacción; climatización a electricidad: nombres de los equipos que requieren puesta a tierra de protección contra rayos y protección antiestática (salas de calderas de gasóleo/gas, chimeneas, tubos de venteo de gas, tuberías de gas combustible, etc.); electricidad a fontanería: requisitos de protección contra incendios de los locales especiales; electricidad a climatización: parámetros de temperatura y humedad y requisitos de ventilación de los locales especiales; sistemas inteligentes a climatización: parámetros de temperatura y humedad y requisitos de ventilación de cada local de corrientes débiles y de la sala de control de seguridad; sistemas inteligentes a electricidad: consumo eléctrico de cada local de corrientes débiles y de la sala de control de seguridad",
        "pt": "Hidráulica para climatização: parâmetros de temperatura e umidade e requisitos de ventilação de ambientes especiais (como os requisitos de temperatura da sala de reservatórios, os nomes dos ambientes de proteção contra incêndio por gás, etc.); climatização para hidráulica: nomes dos ambientes sem aquecimento; climatização para elétrica: nomes dos equipamentos que exigem aterramento contra raios e proteção antiestática (casas de caldeiras a óleo/gás, chaminés, tubos de vente de gás, tubulações de gás combustível, etc.); elétrica para hidráulica: requisitos de proteção contra incêndio de ambientes especiais; elétrica para climatização: parâmetros de temperatura e umidade e requisitos de ventilação de ambientes especiais; sistemas inteligentes para climatização: parâmetros de temperatura e umidade e requisitos de ventilação de cada ambiente de corrente fraca e da sala de controle de segurança; sistemas inteligentes para elétrica: consumo de energia de cada ambiente de corrente fraca e da sala de controle de segurança",
        "ru": "Водоснабжение — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции специальных помещений (например, требования к температуре водяного бака, наименования помещений газового пожаротушения и т.п.) ; отопление и вентиляция — водоснабжению: наименования неотапливаемых помещений; отопление и вентиляция — электроснабжению: наименования оборудования, требующего молниезащитного заземления и антистатической защиты (мазутно-газовые котельные, дымовые трубы, газовые свечи рассеивания, трубопроводы горючих газов и т.п.) ; электроснабжение — водоснабжению: требования пожарной безопасности к специальным помещениям; электроснабжение — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции специальных помещений; интеллектуальные системы — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции слаботочных помещений и диспетчерской охраны; интеллектуальные системы — электроснабжению: электропотребление слаботочных помещений и диспетчерской охраны",
        "ar": "السباكة إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية للغرف الخاصة (مثل متطلبات درجة حرارة غرفة خزان المياه، وأسماء غرف الإطفاء بالغاز وغيرها)؛ والتدفئة والتهوية إلى السباكة: أسماء الغرف غير المدفأة؛ والتدفئة والتهوية إلى الكهرباء: أسماء المعدات التي تتطلب تأريض الحماية من الصواعق والحماية من الكهرباء الساكنة (غرف الغلايات النفطية/الغازية والمداخن وأنابيب تصريف الغاز وخطوط الغاز القابلة للاشتعال وغيرها)؛ والكهرباء إلى السباكة: متطلبات الحماية من الحرائق للغرف الخاصة؛ والكهرباء إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية للغرف الخاصة؛ والأنظمة الذكية إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية لغرف التيارات الضعيفة وغرفة التحكم الأمنية؛ والأنظمة الذكية إلى الكهرباء: استهلاك الكهرباء لغرف التيارات الضعيفة وغرفة التحكم الأمنية",
        "it": "Idrici a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali speciali (come i requisiti di temperatura del locale serbatoi, i nomi dei locali di spegnimento a gas, ecc.); HVAC a idrici: nomi dei locali non riscaldati; HVAC a elettrici: nomi delle apparecchiature che richiedono messa a terra contro i fulmini e protezione antistatica (centrali a olio/gas, camini, tubi di sfogo gas, tubazioni di gas combustibile, ecc.); elettrici a idrici: requisiti di protezione antincendio dei locali speciali; elettrici a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali speciali; sistemi intelligenti a HVAC: parametri di temperatura e umidità e requisiti di ventilazione di ogni locale debole corrente e della sala di controllo sicurezza; sistemi intelligenti a elettrici: consumo elettrico di ogni locale debole corrente e della sala di controllo sicurezza",
        "vi": "Cấp thoát nước gửi điều hòa: thông số nhiệt độ độ ẩm và yêu cầu thông gió của phòng đặc biệt (như yêu cầu nhiệt độ phòng bể nước, tên phòng chữa cháy bằng khí, v.v.); điều hòa gửi cấp thoát nước: tên phòng không sưởi; điều hòa gửi điện: tên thiết bị cần tiếp đất chống sét và chống tĩnh điện (phòng lò hơi dầu khí, ống khói, ống xả khí gas, đường ống khí cháy được, v.v.); điện gửi cấp thoát nước: yêu cầu phòng cháy phòng đặc biệt; điện gửi điều hòa: thông số nhiệt độ độ ẩm và yêu cầu thông gió phòng đặc biệt; hệ thống thông minh gửi điều hòa: thông số nhiệt độ độ ẩm và yêu cầu thông gió của các phòng điện yếu, phòng điều khiển an ninh; hệ thống thông minh gửi điện: lượng điện tiêu thụ của các phòng điện yếu, phòng điều khiển an ninh",
        "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่ทำความร้อน ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องพิเศษ ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณไฟฟ้าที่ใช้ของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง",
    },
    "建筑专业提出：防火分区图；各朝向窗墙比、建筑体型系数等与建筑热工有关的资料": {
        "en": "Proposed by the architectural discipline: fire compartmentation drawings; materials related to building thermal performance such as window-to-wall ratios of each orientation and building shape coefficient",
        "ja": "建築専門が提出：防火区画図；各方位の窓壁比、建築体型係数など建築熱工に関する資料",
        "ko": "건축 전문 제출: 방화 구획도; 각 방위별 창벽비, 건축 체형 계수 등 건축 열공 관련 자료",
        "fr": "Présenté par la discipline architecture : plans de compartimentage coupe-feu ; documents relatifs à la performance thermique du bâtiment tels que les rapports fenêtre/mur de chaque orientation et le coefficient de compacité du bâtiment",
        "de": "Vorgelegt von der Architekturdisziplin: Brandabschnittspläne; Unterlagen zum Wärmeschutz des Gebäudes wie Fenster-Wand-Verhältnisse der einzelnen Himmelsrichtungen und Gebäudeformfaktor",
        "es": "Presentado por la especialidad de arquitectura: planos de compartimentación contra incendios; documentos relativos al comportamiento térmico del edificio, como las relaciones ventana/muro de cada orientación y el coeficiente de forma del edificio",
        "pt": "Apresentado pela especialidade de arquitetura: plantas de compartimentação corta-fogo; documentos relativos ao desempenho térmico do edifício, como as relações janela/parede de cada orientação e o coeficiente de forma do edifício",
        "ru": "Представляется архитектурной специальностью: схемы противопожарных отсеков; материалы, связанные с теплозащитой здания, такие как коэффициенты остекления фасадов по сторонам света и коэффициент компактности здания",
        "ar": "يقدمه تخصص العمارة: مخططات تقسيمات الحماية من الحرائق؛ والمواد المتعلقة بالأداء الحراري للمبنى مثل نسب النوافذ إلى الجدران لكل اتجاه ومعامل شكل المبنى",
        "it": "Presentato dalla disciplina architettonica: elaborati di compartimentazione antincendio; documenti relativi alle prestazioni termiche dell'edificio come i rapporti finestra/parete di ogni orientamento e il coefficiente di compattezza dell'edificio",
        "vi": "Chuyên ngành kiến trúc đưa ra: bản vẽ phân khu chống cháy; tài liệu liên quan đến tính năng nhiệt của công trình như tỷ lệ cửa sổ/tường theo từng hướng và hệ số hình dạng công trình",
        "th": "สาขาสถาปัตยกรรมนำเสนอ: แบบแบ่งเขตอัคคีภัย เอกสารที่เกี่ยวข้องกับสมรรถนะทางความร้อนของอาคาร เช่น อัตราส่วนหน้าต่างต่อผนังแต่ละทิศและค่าสัมประสิทธิ์รูปทรงอาคาร",
    },
    "电专业提给暖专业：电气发热量较大的房间的设备发热量（或估算的电气安装容量），例如变配电室、大型计算机主机房等": {
        "en": "Proposed by the electrical discipline to the HVAC discipline: the heat gain of equipment in rooms with large electrical heat generation (or the estimated electrical installed capacity), such as transformer and distribution rooms, large computer host rooms, etc.",
        "ja": "電専門が暖専門に提出：電気発熱量の大きい部屋の設備発熱量（または推定の電気設備容量）、例：変配電室、大型コンピュータ主機房など",
        "ko": "전기 전문이 난방 전문에게 제출: 전기 발열량이 큰 방의 설비 발열량(또는 추정 전기 설비 용량), 예: 변배전실, 대형 컴퓨터 주기실 등",
        "fr": "Présenté par la discipline électricité à la discipline CVC : l'apport thermique des équipements dans les pièces à forte production de chaleur électrique (ou la puissance installée électrique estimée), telles que les postes de transformation et de distribution, les grandes salles d'hébergement informatique, etc.",
        "de": "Vorgelegt von der Elektrodisziplin an die Disziplin Heizung-Lüftung: die Wärmeabgabe der Anlagen in Räumen mit großer elektrischer Wärmeentwicklung (oder die geschätzte elektrische Anschlussleistung), z. B. Trafo- und Verteilungsräume, große Rechnerhosträume usw.",
        "es": "Presentado por la especialidad de electricidad a la especialidad de climatización: la carga térmica de los equipos en los locales con gran generación de calor eléctrico (o la potencia eléctrica instalada estimada), como salas de transformadores y distribución, grandes salas de ordenadores centrales, etc.",
        "pt": "Apresentado pela especialidade elétrica à especialidade de climatização: a carga térmica dos equipamentos em ambientes com grande geração de calor elétrico (ou a potência instalada elétrica estimada), como salas de transformadores e distribuição, grandes salas de servidores, etc.",
        "ru": "Представляется электротехнической специальностью специальности отопления и вентиляции: тепловыделение оборудования в помещениях с большим электрическим тепловыделением (или расчётная установленная электрическая мощность), например, трансформаторные и распределительные, большие машинные залы и т.п.",
        "ar": "يقدمه تخصص الكهرباء إلى تخصص التدفئة والتهوية: الانبعاث الحراري للمعدات في الغرف ذات التوليد الحراري الكهربائي الكبير (أو القدرة الكهربائية المركبة المقدرة)، مثل غرف المحولات والتوزيع، وغرف الحواسيب الرئيسية الكبيرة وغيرها",
        "it": "Presentato dalla disciplina elettrica alla disciplina HVAC: l'apporto termico delle apparecchiature nei locali con grande generazione di calore elettrico (o la potenza elettrica installata stimata), come sale di trasformazione e distribuzione, grandi sale server, ecc.",
        "vi": "Chuyên ngành điện gửi chuyên ngành điều hòa: lượng tỏa nhiệt của thiết bị trong phòng có tỏa nhiệt điện lớn (hoặc công suất lắp đặt điện ước tính), như phòng trạm biến áp, phòng máy chủ lớn, v.v.",
        "th": "สาขาไฟฟ้าเสนอให้สาขาระบบปรับอากาศ: ความร้อนที่ปล่อยจากอุปกรณ์ในห้องที่มีการปล่อยความร้อนทางไฟฟ้าสูง (หรือกำลังไฟฟ้าติดตั้งโดยประมาณ) เช่น ห้องหม้อแปลงไฟฟ้าและจำหน่าย ห้องเครื่องคอมพิวเตอร์แม่ข่ายขนาดใหญ่ ฯลฯ",
    },
    "水专业提给暖专业：由动力专业提供热源时，提出估算的总热负荷和二次水的温度参数；如果热媒为蒸汽，还需要提出凝结水的回水方式": {
        "en": "Proposed by the plumbing discipline to the HVAC discipline: when the heat source is provided by the power discipline, present the estimated total heat load and the temperature parameters of the secondary water; if the heat medium is steam, also present the condensate return method",
        "ja": "水専門が暖専門に提出：動力専門が熱源を提供する場合、推定総熱負荷と二次水の温度パラメータを提出する；熱媒が蒸気の場合、凝縮水の還水方式も提出する",
        "ko": "급배수 전문이 난방 전문에게 제출: 동력 전문이 열원을 제공할 경우, 추정 총 열부하와 2차수 온도 파라미터를 제출합니다; 열매가 증기일 경우 응축수 회수 방식을 추가로 제출합니다",
        "fr": "Présenté par la discipline plomberie à la discipline CVC : lorsque la source de chaleur est fournie par la discipline énergie, présenter la charge thermique totale estimée et les paramètres de température de l'eau secondaire ; si le fluide caloporteur est la vapeur, présenter également le mode de retour des condensats",
        "de": "Vorgelegt von der Sanitärdisziplin an die Disziplin Heizung-Lüftung: wenn die Wärmequelle von der Energiedisziplin bereitgestellt wird, die geschätzte Gesamtwärmelast und die Temperaturparameter des Sekundärwassers angeben; ist das Wärmemedium Dampf, ist auch die Kondensatrückführung anzugeben",
        "es": "Presentado por la especialidad de fontanería a la especialidad de climatización: cuando la fuente de calor la proporciona la especialidad de energía, presentar la carga térmica total estimada y los parámetros de temperatura del agua secundaria; si el medio calefactor es vapor, presentar también el método de retorno de condensados",
        "pt": "Apresentado pela especialidade de hidráulica à especialidade de climatização: quando a fonte de calor for fornecida pela especialidade de energia, apresentar a carga térmica total estimada e os parâmetros de temperatura da água secundária; se o meio térmico for vapor, apresentar também o método de retorno do condensado",
        "ru": "Представляется специальностью водоснабжения специальности отопления и вентиляции: при предоставлении источника тепла энергетической специальностью представить расчётную суммарную тепловую нагрузку и температурные параметры вторичной воды; если теплоноситель — пар, также представить способ возврата конденсата",
        "ar": "يقدمه تخصص السباكة إلى تخصص التدفئة والتهوية: عند توفير مصدر الحرارة من تخصص الطاقة، تقديم الحمل الحراري الإجمالي المقدر ومعايير درجة حرارة الماء الثانوي؛ وإذا كان وسيط الحرارة بخاراً، فيجب أيضاً تقديم طريقة إرجاع المكثفات",
        "it": "Presentato dalla disciplina idrica alla disciplina HVAC: quando la fonte di calore è fornita dalla disciplina energetica, presentare il carico termico totale stimato e i parametri di temperatura dell'acqua secondaria; se il fluido termovettore è vapore, presentare anche la modalità di ritorno della condensa",
        "vi": "Chuyên ngành cấp thoát nước gửi chuyên ngành điều hòa: khi chuyên ngành năng lượng cấp nguồn nhiệt, nêu tổng tải nhiệt ước tính và thông số nhiệt độ nước thứ cấp; nếu môi chất nhiệt là hơi nước, còn cần nêu phương thức hồi nước ngưng",
        "th": "สาขาประปาเสนอให้สาขาระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ หากตัวกลางความร้อนเป็นไอน้ำ ต้องเสนอวิธีการส่งคืนน้ำควบแน่นด้วย",
    },
    "水提建筑：补充完善确认水泵房、水处理机房、热交换站、热水机房、水箱间、报警阀室等机房面积尺寸净高，集中管道井位置尺寸，水池尺寸与平面定位、地下室集水坑位置尺寸，影响建筑室内布置和外立面的设备（消火栓、冷却塔等）初步位置，最底层垫层厚度要求；暖提建筑：补充完善确认冷冻机房、水泵房、锅炉房、热交换站、燃气调压站（室）等机房面积尺寸净高，各层分散空调通风机房位置面积，集中管道井位置尺寸，锅炉房泄爆井（泄爆面）初步位置面积，地板采暖区域或房间名称及对建筑地面做法厚度要求，室外立面百页窗（新风口、排风口等）初步位置尺寸；暖提水：主要用水点房间名称和估算用水量、主要排水点房间名称和排水性质、冷却水总流量进出水温要求和拟定冷水机组台数；电提建筑：补充完善确认变配电室、发电机房等机房面积尺寸净高（含变配电室夹层要求），各层弱电机房配电小室位置面积，主要强电竖井位置尺寸，楼板厚度要求，发电机冷却进风与排风竖井位置面积；智能化提建筑：补充完善确认进线间、网络机房、电话交换机房、有线电视机房、手机信号覆盖机房、安防控制室、卫星电视等机房面积尺寸净高及特殊机房架空地板要求，弱电竖井大型会议室控制室位置尺寸，楼板厚度要求": {
        "en": "Plumbing to architectural: supplement, improve and confirm the area, dimensions and clear height of machine rooms such as pump rooms, water treatment rooms, heat exchange stations, hot water machine rooms, water tank rooms and alarm valve rooms, the positions and dimensions of centralized pipe shafts, the dimensions and plan positions of pools, the positions and dimensions of basement sumps, the preliminary positions of equipment affecting the interior layout and exterior elevation of the building (fire hydrants, cooling towers, etc.), and the thickness requirements of the lowermost floor screed; HVAC to architectural: supplement, improve and confirm the area, dimensions and clear height of machine rooms such as chiller rooms, pump rooms, boiler rooms, heat exchange stations and gas pressure regulating stations (rooms), the positions and areas of decentralized air-conditioning and ventilation machine rooms on each floor, the positions and dimensions of centralized pipe shafts, the preliminary positions and areas of boiler room explosion vent shafts (vent faces), the names of floor heating areas or rooms and the thickness requirements for building floor construction, and the preliminary positions and dimensions of exterior louver windows (fresh air outlets, exhaust outlets, etc.); HVAC to plumbing: names of rooms with main water consumption points and estimated water consumption, names of rooms with main drainage points and the nature of drainage, the total flow rate of cooling water, inlet and outlet water temperature requirements and the proposed number of chiller units; electrical to architectural: supplement, improve and confirm the area, dimensions and clear height of machine rooms such as transformer and distribution rooms and generator rooms (including the mezzanine requirements of transformer and distribution rooms), the positions and areas of weak-current machine rooms and distribution cubicles on each floor, the positions and dimensions of main strong-current shafts, floor thickness requirements, and the positions and areas of generator cooling air inlet and exhaust shafts; intelligent systems to architectural: supplement, improve and confirm the area, dimensions and clear height of machine rooms such as entrance rooms, network machine rooms, telephone exchange rooms, cable TV rooms, mobile signal coverage rooms, security control rooms and satellite TV rooms, as well as the raised floor requirements of special machine rooms, the positions and dimensions of weak-current shafts, large conference rooms and control rooms, and floor thickness requirements",
        "ja": "水提建築：ポンプ室、水処理室、熱交換ステーション、給湯機室、水槽間、警報弁室などの機械室面積・寸法・天井高、集中パイプシャフトの位置・寸法、水池の寸法と平面位置、地下室集水桝の位置・寸法、建築室内レイアウトと外観立面に影響する設備（消火栓、冷却塔など）の初期位置、最下層敷き均し厚さ要求を補充・改善・確認する；暖提建築：冷凍機室、ポンプ室、ボイラー室、熱交換ステーション、ガス調圧ステーション（室）などの機械室面積・寸法・天井高、各階の分散空調・換気機室の位置・面積、集中パイプシャフトの位置・寸法、ボイラー室の爆発逃がし井（逃がし面）の初期位置・面積、床暖房区域または部屋名と建築床工法・厚さ要求、室外立面ルーバー窓（新風口、排風口など）の初期位置・寸法を補充・改善・確認する；暖提水：主要給水点の部屋名と推定用水量、主要排水点の部屋名と排水性質、冷却水総流量、入口出口水温要求、予定冷水機ユニット台数；電提建築：変配電室、発電機室などの機械室面積・寸法・天井高（変配電室中二階要求を含む）、各階の弱電機室・配電小室の位置・面積、主要強電シャフトの位置・寸法、床版厚さ要求、発電機冷却給気・排気シャフトの位置・面積を補充・改善・確認する；智能化提建築：引込線室、ネットワーク機室、電話交換機室、有線テレビ機室、携帯信号カバレッジ機室、安防制御室、衛星テレビなどの機械室面積・寸法・天井高及び特殊機械室のOA床要求、弱電シャフト・大型会議室・制御室の位置・寸法、床版厚さ要求を補充・改善・確認する",
        "ko": "급배수가 건축에: 펌프실, 수처리실, 열교환 스테이션, 급탕기실, 물탱크실, 경보밸브실 등의 기계실 면적·치수·순높이, 집중 파이프 샤프트 위치·치수, 물탱크 치수와 평면 위치, 지하 집수정 위치·치수, 건축 실내 배치와 외관 입면에 영향을 주는 설비(소화전, 냉각탑 등)의 초기 위치, 최하층 되메움(기초 슬래브) 두께 요구를 보완·개선·확인합니다; 난방이 건축에: 냉동기실, 펌프실, 보일러실, 열교환 스테이션, 가스 압력조정 스테이션(실) 등의 기계실 면적·치수·순높이, 각 층 분산 공조·환기 기계실 위치·면적, 집중 파이프 샤프트 위치·치수, 보일러실 폭발 배출정(배출면) 초기 위치·면적, 바닥난방 구역 또는 실명과 건축 바닥 공법·두께 요구, 외부 입면 루버 창(신풍구, 배풍구 등) 초기 위치·치수를 보완·개선·확인합니다; 난방이 급배수에: 주요 급수점 실명과 추정 용수량, 주요 배수점 실명과 배수 성질, 냉각수 총유량, 출입구 수온 요구, 예정 냉동기 대수; 전기가 건축에: 변배전실, 발전기실 등의 기계실 면적·치수·순높이(변배전실 중간층 요구 포함), 각 층 약전 기계실·배전 소실 위치·면적, 주요 강전 샤프트 위치·치수, 바닥 슬래브 두께 요구, 발전기 냉각 급기·배기 샤프트 위치·면적을 보완·개선·확인합니다; 지능화가 건축에: 인입실, 네트워크 기계실, 전화교환기실, 유선TV 기계실, 휴대폰 신호 커버리지 기계실, 안보 통제실, 위성TV 등의 기계실 면적·치수·순높이 및 특수 기계실 액세스 플로어 요구, 약전 샤프트·대형 회의실·통제실 위치·치수, 바닥 슬래브 두께 요구를 보완·개선·확인합니다",
        "fr": "Plomberie à architecture : compléter, améliorer et confirmer les surfaces, dimensions et hauteurs nettes des locaux techniques tels que les locaux de pompes, les locaux de traitement de l'eau, les postes de livraison de chaleur, les locaux de production d'eau chaude, les locaux de réservoirs et les locaux de vannes d'alarme, les positions et dimensions des gaines techniques centralisées, les dimensions et implantations en plan des bassins, les positions et dimensions des puisards de sous-sol, les positions préliminaires des équipements affectant l'agencement intérieur et l'élévation extérieure du bâtiment (poteaux d'incendie, tours de refroidissement, etc.) et les exigences d'épaisseur de la chape du niveau le plus bas ; CVC à architecture : compléter, améliorer et confirmer les surfaces, dimensions et hauteurs nettes des locaux techniques tels que les locaux des groupes froids, les locaux de pompes, les chaufferies, les postes de livraison de chaleur et les postes (locaux) de détente de gaz, les positions et surfaces des locaux techniques de climatisation et de ventilation décentralisés de chaque niveau, les positions et dimensions des gaines techniques centralisées, les positions et surfaces préliminaires des puits d'évents de soufflage des chaufferies, les noms des zones ou pièces de plancher chauffant et les exigences d'épaisseur des dallages, ainsi que les positions et dimensions préliminaires des fenêtres à persiennes en façade (bouches d'air neuf, bouches d'extraction, etc.) ; CVC à plomberie : noms des pièces des principaux points de consommation d'eau et consommation d'eau estimée, noms des pièces des principaux points de drainage et nature du drainage, débit total de l'eau de refroidissement, exigences de températures d'entrée et de sortie et nombre prévu de groupes refroidisseurs ; électricité à architecture : compléter, améliorer et confirmer les surfaces, dimensions et hauteurs nettes des locaux techniques tels que les postes de transformation et de distribution et les locaux de groupes électrogènes (y compris les exigences de mezzanine des postes de transformation), les positions et surfaces des locaux courants faibles et des petits locaux de distribution de chaque niveau, les positions et dimensions des principales gaines courants forts, les exigences d'épaisseur des planchers et les positions et surfaces des gaines d'air de refroidissement des groupes électrogènes ; systèmes intelligents à architecture : compléter, améliorer et confirmer les surfaces, dimensions et hauteurs nettes des locaux techniques tels que les locaux d'introduction, les locaux de réseaux, les locaux de centraux téléphoniques, les locaux de télévision par câble, les locaux de couverture de signal mobile, les salles de contrôle de sécurité et les locaux de télévision par satellite, ainsi que les exigences de plancher surélevé des locaux techniques spéciaux, les positions et dimensions des gaines courants faibles, des grandes salles de réunion et des salles de contrôle, et les exigences d'épaisseur des planchers",
        "de": "Sanitär an Architektur: die Flächen, Abmessungen und lichten Höhen von Technikräumen wie Pumpenräumen, Wasseraufbereitungsräumen, Wärmeübergabestationen, Warmwasseranlagenräumen, Behälterräumen und Alarmventilräumen, die Positionen und Abmessungen zentraler Installationsschächte, die Abmessungen und Lagepläne von Becken, die Positionen und Abmessungen von Untergeschoss-Sümpfen, die vorläufigen Positionen von Anlagen, die die Innenanordnung und die Außenansicht des Gebäudes beeinflussen (Löschhydranten, Kühltürme usw.), sowie die Dickeanforderungen des untersten Estrichs ergänzen, verbessern und bestätigen; Heizung-Lüftung an Architektur: die Flächen, Abmessungen und lichten Höhen von Technikräumen wie Kältemaschinenräumen, Pumpenräumen, Heizräumen, Wärmeübergabestationen und Gasdruckregelstationen (-räumen), die Positionen und Flächen dezentraler Klima- und Lüftungstechnikräume je Geschoss, die Positionen und Abmessungen zentraler Installationsschächte, die vorläufigen Positionen und Flächen der Explosionsentlastungsschächte von Heizräumen, die Namen von Fußbodenheizungsbereichen oder -räumen und die Anforderungen an die Fußbodenkonstruktion und -dicke des Gebäudes sowie die vorläufigen Positionen und Abmessungen von außenliegenden Lamellenfenstern (Frischluftöffnungen, Abluftöffnungen usw.); Heizung-Lüftung an Sanitär: Namen der Räume mit Hauptwasserverbrauchsstellen und geschätzter Wasserverbrauch, Namen der Räume mit Hauptentwässerungsstellen und Art der Entwässerung, Gesamtdurchfluss des Kühlwassers, Anforderungen an Vor- und Rücklauftemperatur und vorgesehene Anzahl der Kältemaschinen; Elektro an Architektur: die Flächen, Abmessungen und lichten Höhen von Technikräumen wie Trafo- und Verteilungsräumen und Generatorräumen (einschließlich der Anforderungen an Zwischengeschosse der Traforäume), die Positionen und Flächen von Schwachstrom-Technikräumen und Verteilerkabinen je Geschoss, die Positionen und Abmessungen der Hauptstarkstromschächte, die Anforderungen an die Geschossdeckenstärke sowie die Positionen und Flächen der Kühlluftzuluft- und Abluftschächte der Generatoren; Gebäudeautomation an Architektur: die Flächen, Abmessungen und lichten Höhen von Technikräumen wie Eingangsräumen, Netzwerktechnikräumen, Telefonvermittlungsräumen, Kabelfernsehräumen, Räumen für Mobilfunksignalversorgung, Sicherheitskontrollräumen und Satellitenfernsehräumen sowie die Anforderungen an Doppelböden spezieller Technikräume, die Positionen und Abmessungen von Schwachstromschächten, großen Konferenzräumen und Kontrollräumen und die Anforderungen an die Geschossdeckenstärke",
        "es": "Fontanería a arquitectura: complementar, mejorar y confirmar las superficies, dimensiones y alturas libres de los locales técnicos como salas de bombas, salas de tratamiento de agua, estaciones de intercambio de calor, salas de producción de agua caliente, salas de depósitos de agua y salas de válvulas de alarma, las posiciones y dimensiones de los patinillos técnicos centralizados, las dimensiones y emplazamientos en planta de los depósitos, las posiciones y dimensiones de los sumideros del sótano, las posiciones preliminares de los equipos que afectan a la distribución interior y a la fachada del edificio (bocas de incendio, torres de refrigeración, etc.) y los requisitos de espesor de la solera del nivel inferior; climatización a arquitectura: complementar, mejorar y confirmar las superficies, dimensiones y alturas libres de los locales técnicos como salas de enfriadoras, salas de bombas, salas de calderas, estaciones de intercambio de calor y estaciones (locales) de regulación de gas, las posiciones y superficies de los locales técnicos descentralizados de climatización y ventilación de cada planta, las posiciones y dimensiones de los patinillos técnicos centralizados, las posiciones y superficies preliminares de los conductos de alivio de explosión de las salas de calderas, los nombres de las zonas o locales de calefacción por suelo radiante y los requisitos de espesor de los solados, así como las posiciones y dimensiones preliminares de las ventanas de lamas de fachada (tomas de aire nuevo, salidas de extracción, etc.); climatización a fontanería: nombres de los locales de los principales puntos de consumo de agua y consumo de agua estimado, nombres de los locales de los principales puntos de drenaje y naturaleza del drenaje, caudal total del agua de refrigeración, requisitos de temperaturas de entrada y salida y número previsto de grupos frigoríficos; electricidad a arquitectura: complementar, mejorar y confirmar las superficies, dimensiones y alturas libres de los locales técnicos como centros de transformación y salas de grupos electrógenos (incluidos los requisitos de entreplanta de los centros de transformación), las posiciones y superficies de los locales de corrientes débiles y de los cuartos de distribución de cada planta, las posiciones y dimensiones de los patinillos principales de corrientes fuertes, los requisitos de espesor de forjados y las posiciones y superficies de los conductos de entrada y salida de aire de refrigeración de los grupos electrógenos; sistemas inteligentes a arquitectura: complementar, mejorar y confirmar las superficies, dimensiones y alturas libres de los locales técnicos como salas de entrada, salas de redes, salas de centralitas telefónicas, salas de televisión por cable, salas de cobertura de señal móvil, salas de control de seguridad y salas de televisión por satélite, así como los requisitos de suelo técnico elevado de los locales técnicos especiales, las posiciones y dimensiones de los patinillos de corrientes débiles, de las grandes salas de reuniones y de las salas de control, y los requisitos de espesor de forjados",
        "pt": "Hidráulica para arquitetura: complementar, aperfeiçoar e confirmar as áreas, dimensões e pés-direitos dos ambientes técnicos como salas de bombas, salas de tratamento de água, centrais de troca de calor, salas de produção de água quente, salas de reservatórios e salas de válvulas de alarme, as posições e dimensões dos shafts técnicos centralizados, as dimensões e posicionamentos em planta dos reservatórios, as posições e dimensões dos poços de coleta do subsolo, as posições preliminares dos equipamentos que afetam a disposição interna e a fachada do edifício (hidrantes, torres de resfriamento, etc.) e os requisitos de espessura da camada de regularização do nível mais baixo; climatização para arquitetura: complementar, aperfeiçoar e confirmar as áreas, dimensões e pés-direitos dos ambientes técnicos como salas de chillers, salas de bombas, casas de caldeiras, centrais de troca de calor e estações (ambientes) de regulação de gás, as posições e áreas dos ambientes técnicos descentralizados de climatização e ventilação de cada pavimento, as posições e dimensões dos shafts técnicos centralizados, as posições e áreas preliminares dos poços de alívio de explosão das casas de caldeiras, os nomes das áreas ou ambientes de piso radiante e os requisitos de espessura dos revestimentos de piso, bem como as posições e dimensões preliminares das janelas de lâminas na fachada (entradas de ar novo, saídas de exaustão, etc.); climatização para hidráulica: nomes dos ambientes dos principais pontos de consumo de água e consumo de água estimado, nomes dos ambientes dos principais pontos de drenagem e natureza da drenagem, vazão total da água de resfriamento, requisitos de temperaturas de entrada e saída e número previsto de chillers; elétrica para arquitetura: complementar, aperfeiçoar e confirmar as áreas, dimensões e pés-direitos dos ambientes técnicos como subestações e salas de geradores (incluindo os requisitos de mezanino das subestações), as posições e áreas dos ambientes de corrente fraca e pequenas salas de distribuição de cada pavimento, as posições e dimensões dos shafts principais de corrente forte, os requisitos de espessura de lajes e as posições e áreas dos dutos de entrada e saída de ar de refrigeração dos geradores; sistemas inteligentes para arquitetura: complementar, aperfeiçoar e confirmar as áreas, dimensões e pés-direitos dos ambientes técnicos como salas de entrada, salas de redes, salas de centrais telefônicas, salas de TV a cabo, salas de cobertura de sinal móvel, salas de controle de segurança e salas de TV por satélite, bem como os requisitos de piso elevado dos ambientes técnicos especiais, as posições e dimensões dos shafts de corrente fraca, grandes salas de reunião e salas de controle, e os requisitos de espessura de lajes",
        "ru": "Водоснабжение — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений (насосных, водоподготовки, тепловых пунктов, горячего водоснабжения, водомерных, узлов с сигнальными клапанами и т.п.), положение и размеры центральных инженерных шахт, размеры и привязку в плане резервуаров, положение и размеры приямков в подвале, предварительное расположение оборудования, влияющего на внутреннюю планировку и наружный вид здания (пожарные краны, градирни и т.п.), требования к толщине стяжки самого нижнего уровня; отопление и вентиляция — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений (холодильных, насосных, котельных, тепловых пунктов, газорегуляторных пунктов (помещений) и т.п.), расположение и площади децентрализованных венткамер и камер кондиционирования по этажам, положение и размеры центральных инженерных шахт, предварительное расположение и площади взрывных клапанов (разгрузочных поверхностей) котельных, наименования зон или помещений тёплых полов и требования к конструкции и толщине полов, предварительное расположение и размеры фасадных жалюзийных решёток (воздухозаборов, вытяжных отверстий и т.п.); отопление и вентиляция — водоснабжению: наименования помещений основных водоразборных точек и расчётные расходы воды, наименования помещений основных точек водоотведения и характер стоков, общий расход охлаждающей воды, требования к температурам на входе и выходе и предполагаемое число чиллеров; электроснабжение — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений (трансформаторных, дизельных и т.п., включая требования к антресолям трансформаторных), расположение и площади слаботочных и распределительных помещений по этажам, положение и размеры основных сильноточных шахт, требования к толщине перекрытий, положение и площади шахт забора и выброса охлаждающего воздуха дизель-генераторов; интеллектуальные системы — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений (вводных, серверных, АТС, кабельного телевидения, сотовой связи, диспетчерской охраны, спутникового телевидения и т.п.), а также требования к фальшполам специальных помещений, положение и размеры слаботочных шахт, больших конференц-залов и диспетчерских, требования к толщине перекрытий",
        "ar": "السباكة إلى العمارة: استكمال وتحسين وتأكيد مساحات وأبعاد وارتفاعات غرف الآلات مثل غرف المضخات وغرف معالجة المياه ومحطات التبادل الحراري وغرف الماء الساخن وغرف الخزانات وغرف صمامات الإنذار، ومواقع وأبعاد الأعمدة الفنية المركزية، وأبعاد ومواقع الأحواض في المخطط، ومواقع وأبعاد آبار التجميع في الطابق السفلي، والمواقع الأولية للمعدات التي تؤثر على التوزيع الداخلي والواجهة الخارجية للمبنى (صنابير الإطفاء، أبراج التبريد وغيرها)، ومتطلبات سماكة طبقة التسوية في أدنى مستوى؛ والتدفئة والتهوية إلى العمارة: استكمال وتحسين وتأكيد مساحات وأبعاد وارتفاعات غرف الآلات مثل غرف المبردات وغرف المضخات وغرف الغلايات ومحطات التبادل الحراري ومحطات (غرف) تنظيم الغاز، ومواقع ومساحات غرف التكييف والتهوية الموزعة في كل طابق، ومواقع وأبعاد الأعمدة الفنية المركزية، والمواقع والمساحات الأولية لأعمدة تفريغ الانفجار في غرف الغلايات، وأسماء مناطق أو غرف التدفئة الأرضية ومتطلبات سماكة تنفيذ الأرضيات، والمواقع والأبعاد الأولية لنوافذ الشرائح الخارجية (مداخل الهواء الجديد، مخارج الهواء العادم وغيرها)؛ والتدفئة والتهوية إلى السباكة: أسماء غرف نقاط استهلاك المياه الرئيسية واستهلاك المياه التقديري، وأسماء غرف نقاط الصرف الرئيسية وطبيعة الصرف، وإجمالي تدفق مياه التبريد ومتطلبات درجات حرارة الدخول والخروج والعدد المخطط لوحدات التبريد؛ والكهرباء إلى العمارة: استكمال وتحسين وتأكيد مساحات وأبعاد وارتفاعات غرف الآلات مثل غرف المحولات والتوزيع وغرف المولدات (بما في ذلك متطلبات الطابق النصفي لغرف المحولات)، ومواقع ومساحات غرف التيار الضعيف وحجرات التوزيع في كل طابق، ومواقع وأبعاد أعمدة التيار القوي الرئيسية، ومتطلبات سماكة الألواح، ومواقع ومساحات أعمدة دخول وخروج هواء تبريد المولدات؛ والأنظمة الذكية إلى العمارة: استكمال وتحسين وتأكيد مساحات وأبعاد وارتفاعات غرف الآلات مثل غرف الإدخال وغرف الشبكات وغرف مقاسم الهاتف وغرف التلفزيون الكبلي وغرف تغطية إشارة الهاتف المحمول وغرف التحكم الأمنية وغرف التلفزيون الفضائي، بالإضافة إلى متطلبات الأرضيات المرتفعة للغرف الفنية الخاصة، ومواقع وأبعاد أعمدة التيار الضعيف وقاعات الاجتماعات الكبيرة وغرف التحكم، ومتطلبات سماكة الألواح",
        "it": "Idrici a architettura: integrare, migliorare e confermare le aree, le dimensioni e le altezze utili dei locali tecnici come sale pompe, sale trattamento acqua, stazioni di scambio termico, sale produzione acqua calda, sale serbatoi e sale valvole di allarme, le posizioni e dimensioni dei cavedi tecnici centralizzati, le dimensioni e le posizioni planimetriche delle vasche, le posizioni e dimensioni dei pozzetti del seminterrato, le posizioni preliminari delle apparecchiature che influenzano la disposizione interna e l'elevazione esterna dell'edificio (idranti, torri di raffreddamento, ecc.) e i requisiti di spessore del massetto del livello più basso; HVAC a architettura: integrare, migliorare e confermare le aree, le dimensioni e le altezze utili dei locali tecnici come sale gruppi frigoriferi, sale pompe, centrali termiche, stazioni di scambio termico e stazioni (locali) di regolazione del gas, le posizioni e aree dei locali tecnici decentralizzati di climatizzazione e ventilazione di ogni piano, le posizioni e dimensioni dei cavedi tecnici centralizzati, le posizioni e aree preliminari dei camini di sfogo esplosione delle centrali termiche, i nomi delle zone o dei locali a pavimento radiante e i requisiti di spessore delle pavimentazioni, nonché le posizioni e dimensioni preliminari delle finestre a lamelle in facciata (prese d'aria nuova, bocchette di estrazione, ecc.); HVAC a idrici: nomi dei locali dei principali punti di consumo idrico e consumo idrico stimato, nomi dei locali dei principali punti di scarico e natura dello scarico, portata totale dell'acqua di raffreddamento, requisiti di temperatura in ingresso e uscita e numero previsto di gruppi frigoriferi; elettrici a architettura: integrare, migliorare e confermare le aree, le dimensioni e le altezze utili dei locali tecnici come cabine di trasformazione e sale gruppi elettrogeni (inclusi i requisiti di soppalco delle cabine), le posizioni e aree dei locali debole corrente e dei piccoli locali di distribuzione di ogni piano, le posizioni e dimensioni dei cavedi principali di corrente forte, i requisiti di spessore dei solai e le posizioni e aree dei condotti di aspirazione e scarico dell'aria di raffreddamento dei gruppi elettrogeni; sistemi intelligenti a architettura: integrare, migliorare e confermare le aree, le dimensioni e le altezze utili dei locali tecnici come locali ingresso linea, locali rete, locali centralino telefonico, locali TV via cavo, locali copertura segnale mobile, sale controllo sicurezza e locali TV satellitare, nonché i requisiti di pavimento sopraelevato dei locali tecnici speciali, le posizioni e dimensioni dei cavedi debole corrente, delle grandi sale riunioni e delle sale controllo, e i requisiti di spessore dei solai",
        "vi": "Cấp thoát nước gửi kiến trúc: bổ sung, hoàn thiện, xác nhận diện tích, kích thước, chiều cao thông thủy các phòng kỹ thuật như phòng bơm, phòng xử lý nước, trạm trao đổi nhiệt, phòng máy nước nóng, phòng bể nước, phòng van báo động; vị trí, kích thước hộp kỹ thuật tập trung; kích thước và định vị mặt bằng bể nước; vị trí, kích thước hố thu nước tầng hầm; vị trí sơ bộ thiết bị ảnh hưởng bố trí nội thất và mặt đứng công trình (họng cứu hỏa, tháp giải nhiệt, v.v.); yêu cầu chiều dày lớp lót tầng dưới cùng; điều hòa gửi kiến trúc: bổ sung, hoàn thiện, xác nhận diện tích, kích thước, chiều cao thông thủy các phòng kỹ thuật như phòng máy lạnh, phòng bơm, phòng lò hơi, trạm trao đổi nhiệt, trạm (phòng) điều áp gas; vị trí, diện tích các phòng kỹ thuật điều hòa thông gió phân tán từng tầng; vị trí, kích thước hộp kỹ thuật tập trung; vị trí, diện tích sơ bộ giếng xả nổ phòng lò hơi; tên khu vực hoặc phòng sưởi sàn và yêu cầu chiều dày cấu tạo sàn kiến trúc; vị trí, kích thước sơ bộ cửa chớp mặt đứng (miệng gió tươi, miệng gió thải, v.v.); điều hòa gửi cấp thoát nước: tên phòng điểm dùng nước chính và lượng nước dùng ước tính; tên phòng điểm thoát nước chính và tính chất thoát nước; tổng lưu lượng nước làm mát, yêu cầu nhiệt độ nước vào ra và số tổ máy lạnh dự kiến; điện gửi kiến trúc: bổ sung, hoàn thiện, xác nhận diện tích, kích thước, chiều cao thông thủy các phòng kỹ thuật như phòng trạm biến áp, phòng máy phát (kể cả yêu cầu gác lửng phòng trạm biến áp); vị trí, diện tích phòng điện yếu, buồng phân phối từng tầng; vị trí, kích thước hộp kỹ thuật điện mạnh chính; yêu cầu chiều dày sàn; vị trí, diện tích giếng gió vào và gió ra làm mát máy phát; hệ thống thông minh gửi kiến trúc: bổ sung, hoàn thiện, xác nhận diện tích, kích thước, chiều cao thông thủy các phòng kỹ thuật như phòng đầu vào, phòng mạng, phòng tổng đài điện thoại, phòng truyền hình cáp, phòng phủ sóng di động, phòng điều khiển an ninh, phòng truyền hình vệ tinh, cùng yêu cầu sàn nâng cho phòng kỹ thuật đặc biệt; vị trí, kích thước hộp kỹ thuật điện yếu, phòng họp lớn, phòng điều khiển; yêu cầu chiều dày sàn",
        "th": "ประปาส่งให้สถาปัตยกรรม: เพิ่มเติม ปรับปรุง และยืนยันพื้นที่ ขนาด และความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊มน้ำ ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องเครื่องน้ำร้อน ห้องถังน้ำ และห้องวาล์วสัญญาณเตือน ตำแหน่งและขนาดช่องเดินท่อรวม ขนาดและตำแหน่งผังของบ่อน้ำ ตำแหน่งและขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่ส่งผลต่อการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) และข้อกำหนดความหนาของชั้นปรับระดับชั้นล่างสุด ระบบปรับอากาศส่งให้สถาปัตยกรรม: เพิ่มเติม ปรับปรุง และยืนยันพื้นที่ ขนาด และความสูงโปร่งของห้องเครื่อง เช่น ห้องเครื่องทำความเย็น ห้องปั๊มน้ำ ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน และสถานี (ห้อง) ควบคุมแรงดันแก๊ส ตำแหน่งและพื้นที่ห้องเครื่องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งและขนาดช่องเดินท่อรวม ตำแหน่งและพื้นที่เบื้องต้นของช่องระบายแรงดันระเบิดห้องหม้อไอน้ำ ชื่อพื้นที่หรือห้องทำความร้อนพื้นและข้อกำหนดความหนาของพื้นอาคาร และตำแหน่งและขนาดเบื้องต้นของช่องระบายอากาศแบบใบพัดภายนอก (ช่องลมเข้า ช่องลมออก ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้าและออก และจำนวนเครื่องทำความเย็นที่กำหนด ไฟฟ้าส่งให้สถาปัตยกรรม: เพิ่มเติม ปรับปรุง และยืนยันพื้นที่ ขนาด และความสูงโปร่งของห้องเครื่อง เช่น ห้องหม้อแปลงและจ่ายไฟฟ้า และห้องเครื่องกำเนิดไฟฟ้า (รวมข้อกำหนดชั้นลอยของห้องหม้อแปลง) ตำแหน่งและพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งและขนาดช่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาของพื้น และตำแหน่งและพื้นที่ช่องอากาศเข้าและออกเพื่อระบายความร้อนเครื่องกำเนิดไฟฟ้า ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เพิ่มเติม ปรับปรุง และยืนยันพื้นที่ ขนาด และความสูงโปร่งของห้องเครื่อง เช่น ห้องอินพุต ห้องเครือข่าย ห้องตู้สลับโทรศัพท์ ห้องทีวีเคเบิล ห้องครอบคลุมสัญญาณโทรศัพท์มือถือ ห้องควบคุมความปลอดภัย และห้องทีวีดาวเทียม รวมถึงข้อกำหนดพื้นยกสูงของห้องเครื่องพิเศษ ตำแหน่งและขนาดช่องไฟฟ้าอ่อน ห้องประชุมขนาดใหญ่ และห้องควบคุม และข้อกำหนดความหนาของพื้น",
    },
    "水、暖、电、智能化提结构：位于楼板上的机房设备和冷却塔的结构荷载；最大长边尺寸或直径≥1000mm的承重结构墙和楼板的预留洞初步位置；水、暖提电：各强电设备的位置、电压等级及其电气安装容量（消防设备和平时使用的设备分别提出）": {
        "en": "Plumbing, HVAC, electrical and intelligent systems to structural: the structural loads of machine room equipment and cooling towers located on floor slabs; the preliminary positions of reserved openings with maximum long-side dimension or diameter ≥1000mm in load-bearing structural walls and floor slabs; plumbing and HVAC to electrical: the positions, voltage levels and electrical installed capacities of various strong-current equipment (fire protection equipment and equipment for normal use proposed separately)",
        "ja": "水・暖・電・智能化提構造：床版上の機室設備と冷却塔の構造荷重；最大長辺寸法または直径≥1000mmの耐力構造壁と床版の予備開口初期位置；水・暖提電：各強電設備の位置、電圧レベル及びその電気設備容量（消防設備と平時使用設備は別々に提出）",
        "ko": "급배수·난방·전기·지능화가 구조에: 바닥 슬래브 위 기계실 설비와 냉각탑의 구조 하중; 최대 장변 치수 또는 직경 ≥1000mm의 내력 구조벽과 슬래브 예비 구멍의 초기 위치; 급배수·난방이 전기에: 각 강전 설비의 위치, 전압 레벨 및 전기 설비 용량(소방 설비와 평시 사용 설비는 별도 제출)",
        "fr": "Plomberie, CVC, électricité et systèmes intelligents à structure : les charges structurelles des équipements de locaux techniques et des tours de refroidissement situés sur les dalles ; les positions préliminaires des réservations dont le plus grand côté ou diamètre est ≥ 1000 mm dans les murs porteurs et les dalles ; plomberie et CVC à électricité : les positions, niveaux de tension et puissances installées électriques des divers équipements courants forts (les équipements de protection incendie et ceux d'usage courant proposés séparément)",
        "de": "Sanitär, Heizung-Lüftung, Elektro und Gebäudeautomation an Tragwerk: die Tragwerkslasten der auf Geschossdecken befindlichen Technikraumanlagen und Kühltürme; die vorläufigen Positionen von Durchbrüchen mit maximaler Längsseite oder Durchmesser ≥ 1000 mm in tragenden Wänden und Geschossdecken; Sanitär und Heizung-Lüftung an Elektro: die Positionen, Spannungsebenen und elektrischen Anschlussleistungen der verschiedenen Starkstromanlagen (Brandschutzanlagen und Anlagen für den Normalbetrieb getrennt angeben)",
        "es": "Fontanería, climatización, electricidad y sistemas inteligentes a estructuras: las cargas estructurales de los equipos de los locales técnicos y de las torres de refrigeración situados sobre los forjados; las posiciones preliminares de los huecos cuyo lado mayor o diámetro sea ≥ 1000 mm en muros portantes y forjados; fontanería y climatización a electricidad: las posiciones, niveles de tensión y potencias eléctricas instaladas de los diversos equipos de corrientes fuertes (los equipos de protección contra incendios y los de uso normal se proponen por separado)",
        "pt": "Hidráulica, climatização, elétrica e sistemas inteligentes para estruturas: as cargas estruturais dos equipamentos de ambientes técnicos e das torres de resfriamento localizados sobre lajes; as posições preliminares das aberturas com maior lado ou diâmetro ≥ 1000 mm em paredes estruturais e lajes; hidráulica e climatização para elétrica: as posições, níveis de tensão e potências instaladas elétricas dos diversos equipamentos de corrente forte (equipamentos de proteção contra incêndio e de uso normal propostos separadamente)",
        "ru": "Водоснабжение, отопление и вентиляция, электроснабжение и интеллектуальные системы — конструкциям: нагрузки на конструкции от оборудования технических помещений и градирен, расположенных на перекрытиях; предварительное расположение отверстий с наибольшим размером стороны или диаметром ≥ 1000 мм в несущих стенах и плитах перекрытий; водоснабжение и отопление и вентиляция — электроснабжению: расположение, уровни напряжения и установленные электрические мощности различных сильноточных устройств (противопожарное оборудование и оборудование повседневного использования указываются отдельно)",
        "ar": "السباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية إلى الإنشاءات: الأحمال الإنشائية لمعدات غرف الآلات وأبراج التبريد الموجودة على الألواح؛ والمواقع الأولية للفتحات التي يبلغ أكبر بعد أو قطر لها ≥ 1000 مم في الجدران الإنشائية الحاملة والألواح؛ والسباكة والتدفئة والتهوية إلى الكهرباء: مواقع ومستويات الجهد والقدرات الكهربائية المركبة لمختلف معدات التيار القوي (يتم اقتراح معدات الحماية من الحرائق ومعدات الاستخدام العادي بشكل منفصل)",
        "it": "Idrici, HVAC, elettrici e sistemi intelligenti a strutture: i carichi strutturali delle apparecchiature dei locali tecnici e delle torri di raffreddamento situate sui solai; le posizioni preliminari delle aperture con lato maggiore o diametro ≥ 1000 mm nelle pareti strutturali portanti e nei solai; idrici e HVAC a elettrici: le posizioni, i livelli di tensione e le potenze elettriche installate delle varie apparecchiature di corrente forte (le apparecchiature antincendio e quelle di uso ordinario proposte separatamente)",
        "vi": "Cấp thoát nước, điều hòa, điện, hệ thống thông minh gửi kết cấu: tải trọng kết cấu của thiết bị phòng kỹ thuật và tháp giải nhiệt đặt trên sàn; vị trí sơ bộ lỗ chờ có cạnh dài lớn nhất hoặc đường kính ≥1000mm trên tường chịu lực và sàn; cấp thoát nước, điều hòa gửi điện: vị trí, cấp điện áp và công suất lắp đặt điện của các thiết bị điện mạnh (thiết bị phòng cháy và thiết bị sử dụng thường xuyên nêu riêng)",
        "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: น้ำหนักบรรทุกโครงสร้างของอุปกรณ์ห้องเครื่องและหอหล่อเย็นที่อยู่บนพื้น ตำแหน่งเบื้องต้นของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนักและพื้น ประปาและระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่ง ระดับแรงดันไฟฟ้า และกำลังไฟฟ้าติดตั้งของอุปกรณ์ไฟฟ้าแรงต่าง ๆ (อุปกรณ์ดับเพลิงและอุปกรณ์ใช้ปกติเสนอแยกกัน)",
    },
    "总图专业提出（在第一阶段基础上）：含规划用地坐标图在内的各建筑单体轴线或外墙定位图；各单体建筑±0.000标高与绝对标高的关系；红线内已有建筑、构筑物、道路、绿化等的名称、布置与定位；市政管线走向及平面位置；其他影响设计的资料（排洪沟、保留市政管线等）。接受专业：建筑、结构、水、暖、电、经济": {
        "en": "Proposed by the site planning discipline (on the basis of the first stage): the axis or exterior wall positioning drawings of each building block including the planning land coordinate drawing; the relationship between the ±0.000 elevation of each building block and the absolute elevation; the names, layout and positioning of existing buildings, structures, roads and greenery within the red line; the alignment and plan positions of municipal pipelines; other information affecting the design (flood drainage ditches, retained municipal pipelines, etc.). Receiving disciplines: architecture, structural, plumbing, HVAC, electrical, economics",
        "ja": "総図専門が提出（第一段階に基づく）：計画用地座標図を含む各建築棟の軸線または外壁位置図；各建築棟の±0.000標高と絶対標高の関係；境界線内の既存建築、構築物、道路、緑化などの名称、配置と位置；市政配管の走向と平面位置；設計に影響するその他の資料（排洪溝、保留市政配管など）。受領専門：建築、構造、水、暖、電、経済",
        "ko": "총도 전문이 제출(제1단계 기반): 계획 용지 좌표도를 포함한 각 건축 동의 축선 또는 외벽 위치도; 각 건축 동의 ±0.000 표고와 절대 표고의 관계; 경계선 내 기존 건축물, 구조물, 도로, 조경 등의 명칭, 배치와 위치; 시정 배관의 방향과 평면 위치; 설계에 영향을 주는 기타 자료(배홍구, 보존 시정 배관 등). 수령 전문: 건축, 구조, 급배수, 난방, 전기, 경제",
        "fr": "Présenté par la discipline aménagement du site (sur la base de la première phase) : les plans de positionnement des axes ou des façades de chaque bâtiment, y compris le plan de coordonnées du terrain ; la relation entre la cote ±0.000 de chaque bâtiment et l'altitude absolue ; les noms, dispositions et implantations des bâtiments, ouvrages, voiries et espaces verts existants à l'intérieur de la limite ; les tracés et positions en plan des réseaux municipaux ; les autres documents affectant la conception (fossés d'évacuation des crues, réseaux municipaux conservés, etc.). Disciplines réceptrices : architecture, structure, plomberie, CVC, électricité, économie",
        "de": "Vorgelegt von der Standortplanungsdisziplin (auf der Grundlage der ersten Phase): die Achs- oder Außenwand-Positionspläne jedes Gebäudeteils einschließlich des Koordinatenplans des Baugrundstücks; die Beziehung zwischen der ±0.000-Höhe jedes Gebäudeteils und der absoluten Höhe; die Namen, Anordnung und Positionierung bestehender Gebäude, Bauwerke, Straßen und Grünanlagen innerhalb der roten Linie; der Verlauf und die Planpositionen der städtischen Leitungen; weitere entwurfsrelevante Unterlagen (Hochwasserabzugsgräben, erhaltene städtische Leitungen usw.). Empfangende Disziplinen: Architektur, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de urbanización (sobre la base de la primera fase): los planos de localización de ejes o fachadas de cada edificio, incluido el plano de coordenadas del terreno; la relación entre la cota ±0.000 de cada edificio y la altitud absoluta; los nombres, disposición y ubicación de los edificios, obras, viales y zonas verdes existentes dentro del límite; los trazados y posiciones en planta de las redes municipales; otros documentos que afectan al diseño (cauces de evacuación de crecidas, redes municipales conservadas, etc.). Especialidades receptoras: arquitectura, estructuras, fontanería, climatización, electricidad, economía",
        "pt": "Apresentado pela especialidade de urbanismo (com base na primeira fase): as plantas de posicionamento de eixos ou fachadas de cada bloco, incluindo a planta de coordenadas do terreno; a relação entre a cota ±0.000 de cada bloco e a altitude absoluta; os nomes, disposições e posicionamentos dos edifícios, obras, vias e áreas verdes existentes dentro do limite; os traçados e posições em planta das redes municipais; outros documentos que afetam o projeto (canaletas de drenagem de cheias, redes municipais preservadas, etc.). Especialidades receptoras: arquitetura, estruturas, hidráulica, climatização, elétrica, economia",
        "ru": "Представляется специальностью генплана (на основе первой стадии): планы осей или наружных стен каждого здания, включая план с координатами территории; соотношение отметки ±0.000 каждого здания с абсолютной отметкой; наименования, расположение и привязка существующих зданий, сооружений, дорог и озеленения в пределах красной линии; трассы и положение в плане городских сетей; прочие материалы, влияющие на проектирование (водоотводные канавы, сохраняемые городские сети и т.п.). Получающие специальности: архитектура, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, экономика",
        "ar": "يقدمه تخصص تخطيط المواقع (على أساس المرحلة الأولى): مخططات تحديد محاور أو جدران خارجية لكل مبنى بما في ذلك مخطط إحداثيات الأرض المخططة؛ والعلاقة بين منسوب ±0.000 لكل مبنى والمنسوب المطلق؛ وأسماء وتوزيع ومواقع المباني والمنشآت والطرق والمساحات الخضراء القائمة داخل خط الحدود؛ واتجاهات ومواقع شبكات المرافق البلدية في المخطط؛ وغيرها من المعلومات المؤثرة على التصميم (قنوات تصريف الفيضانات، شبكات المرافق المحتفظ بها وغيرها). التخصصات المستلمة: العمارة والإنشاءات والسباكة والتدفئة والتهوية والكهرباء والاقتصاد",
        "it": "Presentato dalla disciplina di pianificazione del sito (sulla base della prima fase): gli elaborati di posizionamento degli assi o delle pareti esterne di ogni edificio, incluso il piano con le coordinate del terreno; il rapporto tra la quota ±0.000 di ogni edificio e la quota assoluta; i nomi, le disposizioni e i posizionamenti di edifici, opere, strade e aree verdi esistenti entro il limite; i tracciati e le posizioni planimetriche delle reti comunali; gli altri documenti che influenzano la progettazione (canali di sfioro, reti comunali conservate, ecc.). Discipline riceventi: architettura, strutture, idrici, HVAC, elettrici, economia",
        "vi": "Chuyên ngành tổng mặt bằng đưa ra (trên cơ sở giai đoạn 1): bản vẽ định vị trục hoặc tường ngoài của từng khối nhà, kể cả bản vẽ tọa độ đất quy hoạch; quan hệ giữa cao độ ±0.000 của từng khối nhà và cao độ tuyệt đối; tên, bố trí và định vị các công trình, kết cấu, đường, cây xanh hiện hữu trong ranh giới; hướng tuyến và vị trí mặt bằng đường ống hạ tầng; các tài liệu khác ảnh hưởng thiết kế (mương thoát lũ, đường ống hạ tầng giữ lại, v.v.). Chuyên ngành nhận: kiến trúc, kết cấu, cấp thoát nước, điều hòa, điện, kinh tế",
        "th": "สาขาวางผังบริเวณนำเสนอ (บนพื้นฐานระยะแรก): แบบกำหนดตำแหน่งแกนหรือผนังภายนอกของแต่ละอาคาร รวมถึงแบบพิกัดที่ดินวางแผน ความสัมพันธ์ระหว่างระดับ ±0.000 ของแต่ละอาคารกับระดับสัมบูรณ์ ชื่อ การจัดวาง และตำแหน่งของอาคาร โครงสร้าง ถนน และพื้นที่สีเขียวที่มีอยู่ภายในเส้นเขต แนวทางและตำแหน่งผังของท่อสาธารณูปโภค เอกสารอื่นที่ส่งผลต่อการออกแบบ (ร่องระบายน้ำหลาก ท่อสาธารณูปโภคที่เก็บรักษา ฯลฯ) สาขาที่รับ: สถาปัตยกรรม โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า เศรษฐศาสตร์",
    },
    "结构专业提出：楼面、屋面结构梁板平面图；与建筑有联系的挡土墙结构形式（提建筑、总图专业）；设计配合所必需的特殊结构构造做法或详图（提建筑、总图专业）。接受专业：建筑、水、暖、电、智能化、经济": {
        "en": "Proposed by the structural discipline: structural beam and slab plans of floors and roofs; the structural forms of retaining walls related to architecture (provided to the architectural and site planning disciplines); special structural construction practices or details necessary for design coordination (provided to the architectural and site planning disciplines). Receiving disciplines: architecture, plumbing, HVAC, electrical, intelligent systems, economics",
        "ja": "構造専門が提出：床・屋根の構造梁板平面図；建築と関連する擁壁の構造形式（建築、総図専門へ提出）；設計配合に必要な特殊構造の工法または詳細図（建築、総図専門へ提出）。受領専門：建築、水、暖、電、智能化、経済",
        "ko": "구조 전문이 제출: 바닥·지붕 구조 보-슬래브 평면도; 건축과 관련된 옹벽 구조 형식(건축, 총도 전문에 제출); 설계 협력에 필요한 특수 구조 시공법 또는 상세도(건축, 총도 전문에 제출). 수령 전문: 건축, 급배수, 난방, 전기, 지능화, 경제",
        "fr": "Présenté par la discipline structure : les plans de structure des planchers et des toitures (poutres et dalles) ; les formes structurelles des murs de soutènement en lien avec l'architecture (fournies aux disciplines architecture et aménagement du site) ; les pratiques constructives ou détails structurels spéciaux nécessaires à la coordination de la conception (fournis aux disciplines architecture et aménagement du site). Disciplines réceptrices : architecture, plomberie, CVC, électricité, systèmes intelligents, économie",
        "de": "Vorgelegt von der Tragwerksdisziplin: Strukturpläne der Decken- und Dachträger und -platten; die mit der Architektur verbundenen Stützmauer-Strukturformen (an die Architektur- und Standortplanungsdisziplin); besondere Konstruktionsweisen oder Details, die für die Entwurfskoordination erforderlich sind (an die Architektur- und Standortplanungsdisziplin). Empfangende Disziplinen: Architektur, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de estructuras: los planos de estructura de forjados y cubiertas (vigas y losas); las formas estructurales de los muros de contención relacionados con la arquitectura (aportados a las especialidades de arquitectura y urbanización); las prácticas constructivas o detalles estructurales especiales necesarios para la coordinación del diseño (aportados a las especialidades de arquitectura y urbanización). Especialidades receptoras: arquitectura, fontanería, climatización, electricidad, sistemas inteligentes, economía",
        "pt": "Apresentado pela especialidade de estruturas: as plantas estruturais de vigas e lajes de pisos e coberturas; as formas estruturais dos muros de arrimo relacionados à arquitetura (fornecidas às especialidades de arquitetura e urbanismo); as práticas construtivas ou detalhes estruturais especiais necessários à coordenação do projeto (fornecidos às especialidades de arquitetura e urbanismo). Especialidades receptoras: arquitetura, hidráulica, climatização, elétrica, sistemas inteligentes, economia",
        "ru": "Представляется конструкторской специальностью: планы несущих балок и плит перекрытий и покрытий; конструктивные формы подпорных стен, связанных с архитектурой (передаются архитектурной специальности и специальности генплана); особые конструктивные решения или детали, необходимые для координации проектирования (передаются архитектурной специальности и специальности генплана). Получающие специальности: архитектура, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы, экономика",
        "ar": "يقدمه تخصص الإنشاءات: مخططات الحزم والألواح الإنشائية للأرضيات والأسقف؛ والأشكال الإنشائية للجدران الاستنادية المرتبطة بالعمارة (تُقدم إلى تخصصي العمارة وتخطيط المواقع)؛ وممارسات البناء الإنشائية الخاصة أو التفاصيل اللازمة لتنسيق التصميم (تُقدم إلى تخصصي العمارة وتخطيط المواقع). التخصصات المستلمة: العمارة والسباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية والاقتصاد",
        "it": "Presentato dalla disciplina strutturale: le piante strutturali di travi e solai di pavimenti e coperture; le forme strutturali dei muri di contenimento collegati all'architettura (fornite alle discipline architettonica e di pianificazione del sito); le pratiche costruttive speciali o i dettagli strutturali necessari per il coordinamento progettuale (forniti alle discipline architettonica e di pianificazione del sito). Discipline riceventi: architettura, idrici, HVAC, elettrici, sistemi intelligenti, economia",
        "vi": "Chuyên ngành kết cấu đưa ra: bản vẽ mặt bằng kết cấu dầm sàn tầng và mái; hình thức kết cấu tường chắn liên quan đến kiến trúc (gửi chuyên ngành kiến trúc, tổng mặt bằng); cấu tạo hoặc chi tiết kết cấu đặc biệt cần thiết cho phối hợp thiết kế (gửi chuyên ngành kiến trúc, tổng mặt bằng). Chuyên ngành nhận: kiến trúc, cấp thoát nước, điều hòa, điện, hệ thống thông minh, kinh tế",
        "th": "สาขาโครงสร้างนำเสนอ: แผนผังคานและพื้นโครงสร้างของพื้นและหลังคา รูปแบบโครงสร้างกำแพงกันดินที่เกี่ยวข้องกับสถาปัตยกรรม (ส่งให้สาขาสถาปัตยกรรมและวางผังบริเวณ) วิธีการก่อสร้างโครงสร้างพิเศษหรือรายละเอียดที่จำเป็นสำหรับการประสานการออกแบบ (ส่งให้สาขาสถาปัตยกรรมและวางผังบริเวณ) สาขาที่รับ: สถาปัตยกรรม ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ เศรษฐศาสตร์",
    },
    "水提总图：各市政管网（给水、排水、热水等）至本楼接口大致位置，主要室外管道布置及主要构筑物（化粪池、隔油池、水表井、水泵接合器等）大致位置；暖提总图：市政管网（热力、燃气）至本楼接口大致位置，主要室外管道布置及构筑物（阀门井等）大致位置；电、智能化提总图：市政管线（电力、通讯、有线电视等）至本楼接口大致位置，主要室外管线布置及构筑物（检修井等）大致位置": {
        "en": "Plumbing to site planning: approximate positions of the interfaces of various municipal pipeline networks (water supply, drainage, hot water, etc.) to this building, approximate positions of major outdoor pipeline layouts and main structures (septic tanks, grease traps, water meter wells, fire pump connections, etc.); HVAC to site planning: approximate positions of the interfaces of municipal pipeline networks (heating, gas) to this building, approximate positions of major outdoor pipeline layouts and structures (valve wells, etc.); electrical and intelligent systems to site planning: approximate positions of the interfaces of municipal pipelines (power, communications, cable TV, etc.) to this building, approximate positions of major outdoor pipeline layouts and structures (inspection wells, etc.)",
        "ja": "水提総図：各市政管網（給水、排水、給湯など）から本棟への接続おおよその位置、主要屋外配管配置及び主要構築物（化糞池、グリーストラップ、水道メーター井、ポンプ接合器など）のおおよその位置；暖提総図：市政管網（熱力、ガス）から本棟への接続おおよその位置、主要屋外配管配置及び構築物（バルブ井など）のおおよその位置；電・智能化提総図：市政管線（電力、通信、有線テレビなど）から本棟への接続おおよその位置、主要屋外配管配置及び構築物（点検井など）のおおよその位置",
        "ko": "급배수가 총도에: 각 시정 관망(급수, 배수, 온수 등)의 본 건물 접속 대략 위치, 주요 옥외 배관 배치 및 주요 구조물(정화조, 유분리조, 수도계량기 우물, 펌프 접합기 등) 대략 위치; 난방이 총도에: 시정 관망(열력, 가스)의 본 건물 접속 대략 위치, 주요 옥외 배관 배치 및 구조물(밸브 우물 등) 대략 위치; 전기·지능화가 총도에: 시정 관선(전력, 통신, 유선TV 등)의 본 건물 접속 대략 위치, 주요 옥외 배관 배치 및 구조물(점검 우물 등) 대략 위치",
        "fr": "Plomberie à l'aménagement du site : positions approximatives des interfaces des divers réseaux municipaux (adduction, assainissement, eau chaude, etc.) avec ce bâtiment, positions approximatives des principaux réseaux extérieurs et ouvrages (fosses septiques, séparateurs à graisses, regards de compteurs d'eau, raccordements de pompes d'incendie, etc.) ; CVC à l'aménagement du site : positions approximatives des interfaces des réseaux municipaux (chauffage, gaz) avec ce bâtiment, positions approximatives des principaux réseaux extérieurs et ouvrages (regards de vannes, etc.) ; électricité et systèmes intelligents à l'aménagement du site : positions approximatives des interfaces des réseaux municipaux (électricité, télécommunications, télévision par câble, etc.) avec ce bâtiment, positions approximatives des principaux réseaux extérieurs et ouvrages (regards de visite, etc.)",
        "de": "Sanitär an Standortplanung: ungefähre Positionen der Anschlusspunkte der verschiedenen städtischen Netze (Wasserversorgung, Entwässerung, Warmwasser usw.) an dieses Gebäude, ungefähre Positionen der Hauptaußenleitungen und der wichtigsten Bauwerke (Kläranlagen, Fettabscheider, Wasserzählerschächte, Feuerlöschpumpenanschlüsse usw.); Heizung-Lüftung an Standortplanung: ungefähre Positionen der Anschlusspunkte der städtischen Netze (Fernwärme, Gas) an dieses Gebäude, ungefähre Positionen der Hauptaußenleitungen und Bauwerke (Absperrschieberschächte usw.); Elektro und Gebäudeautomation an Standortplanung: ungefähre Positionen der Anschlusspunkte der städtischen Leitungen (Strom, Kommunikation, Kabelfernsehen usw.) an dieses Gebäude, ungefähre Positionen der Hauptaußenleitungen und Bauwerke (Inspektionsschächte usw.)",
        "es": "Fontanería a urbanización: posiciones aproximadas de los puntos de conexión de las distintas redes municipales (abastecimiento, saneamiento, agua caliente, etc.) con este edificio, posiciones aproximadas de las principales redes exteriores y obras (fosas sépticas, separadores de grasas, arquetas de contadores de agua, conexiones de bombas de incendio, etc.); climatización a urbanización: posiciones aproximadas de los puntos de conexión de las redes municipales (calefacción, gas) con este edificio, posiciones aproximadas de las principales redes exteriores y obras (arquetas de válvulas, etc.); electricidad y sistemas inteligentes a urbanización: posiciones aproximadas de los puntos de conexión de las redes municipales (electricidad, comunicaciones, televisión por cable, etc.) con este edificio, posiciones aproximadas de las principales redes exteriores y obras (arquetas de registro, etc.)",
        "pt": "Hidráulica para urbanismo: posições aproximadas dos pontos de conexão das diversas redes municipais (abastecimento, drenagem, água quente, etc.) com este edifício, posições aproximadas dos principais traçados externos e obras (fossas sépticas, separadores de gordura, caixas de hidrômetros, conexões de bombas de incêndio, etc.); climatização para urbanismo: posições aproximadas dos pontos de conexão das redes municipais (aquecimento, gás) com este edifício, posições aproximadas dos principais traçados externos e obras (caixas de válvulas, etc.); elétrica e sistemas inteligentes para urbanismo: posições aproximadas dos pontos de conexão das redes municipais (energia, comunicações, TV a cabo, etc.) com este edifício, posições aproximadas dos principais traçados externos e obras (poços de inspeção, etc.)",
        "ru": "Водоснабжение — генплану: приблизительное расположение точек присоединения городских сетей (водоснабжение, водоотведение, горячее водоснабжение и т.п.) к данному зданию, приблизительное расположение основных наружных сетей и сооружений (септики, жироуловители, колодцы водомеров, пожарные соединительные головки и т.п.); отопление и вентиляция — генплану: приблизительное расположение точек присоединения городских сетей (теплоснабжение, газ) к данному зданию, приблизительное расположение основных наружных сетей и сооружений (задвижки в колодцах и т.п.); электроснабжение и интеллектуальные системы — генплану: приблизительное расположение точек присоединения городских линий (электроснабжение, связь, кабельное телевидение и т.п.) к данному зданию, приблизительное расположение основных наружных линий и сооружений (смотровые колодцы и т.п.)",
        "ar": "السباكة إلى تخطيط المواقع: المواقع التقريبية لنقاط ربط شبكات المرافق البلدية المختلفة (إمداد المياه، الصرف، الماء الساخن وغيرها) بهذا المبنى، والمواقع التقريبية للشبكات الخارجية الرئيسية والمنشآت الرئيسية (خزانات الصرف الصحي، فواصل الدهون، آبار عدادات المياه، وصلات مضخات الإطفاء وغيرها)؛ والتدفئة والتهوية إلى تخطيط المواقع: المواقع التقريبية لنقاط ربط شبكات المرافق البلدية (الحرارة، الغاز) بهذا المبنى، والمواقع التقريبية للشبكات الخارجية الرئيسية والمنشآت (آبار الصمامات وغيرها)؛ والكهرباء والأنظمة الذكية إلى تخطيط المواقع: المواقع التقريبية لنقاط ربط خطوط المرافق البلدية (الكهرباء، الاتصالات، التلفزيون الكبلي وغيرها) بهذا المبنى، والمواقع التقريبية للخطوط الخارجية الرئيسية والمنشآت (آبار التفتيش وغيرها)",
        "it": "Idrici a pianificazione del sito: posizioni approssimative dei punti di collegamento delle varie reti comunali (acquedotto, fognatura, acqua calda, ecc.) con questo edificio, posizioni approssimative dei principali tracciati esterni e delle opere (fosse settiche, separatori di grassi, pozzetti contatori acqua, attacchi pompe antincendio, ecc.); HVAC a pianificazione del sito: posizioni approssimative dei punti di collegamento delle reti comunali (teleriscaldamento, gas) con questo edificio, posizioni approssimative dei principali tracciati esterni e delle opere (pozzetti valvole, ecc.); elettrici e sistemi intelligenti a pianificazione del sito: posizioni approssimative dei punti di collegamento delle linee comunali (energia, comunicazioni, TV via cavo, ecc.) con questo edificio, posizioni approssimative dei principali tracciati esterni e delle opere (pozzetti di ispezione, ecc.)",
        "vi": "Cấp thoát nước gửi tổng mặt bằng: vị trí tương đối các điểm đấu nối của các mạng hạ tầng (cấp nước, thoát nước, nước nóng, v.v.) vào công trình này, vị trí tương đối bố trí đường ống ngoài nhà chính và công trình chính (bể tự hoại, bể tách dầu mỡ, giếng đồng hồ nước, đầu nối máy bơm chữa cháy, v.v.); điều hòa gửi tổng mặt bằng: vị trí tương đối điểm đấu nối mạng hạ tầng (nhiệt, gas) vào công trình, vị trí tương đối bố trí đường ống ngoài nhà chính và công trình (giếng van, v.v.); điện, hệ thống thông minh gửi tổng mặt bằng: vị trí tương đối điểm đấu nối đường ống hạ tầng (điện, viễn thông, truyền hình cáp, v.v.) vào công trình, vị trí tương đối bố trí đường ống ngoài chính và công trình (giếng kiểm tra, v.v.)",
        "th": "ประปาส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อเครือข่ายสาธารณูปโภคต่าง ๆ (ประปา ท่อระบายน้ำ น้ำร้อน ฯลฯ) เข้าสู่อาคารนี้ ตำแหน่งโดยประมาณของการวางท่อภายนอกหลักและโครงสร้างหลัก (ถังบำบัด ถังดักไขมัน บ่อมาตรวัดน้ำ ข้อต่อปั๊มดับเพลิง ฯลฯ) ระบบปรับอากาศส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อเครือข่ายสาธารณูปโภค (ความร้อน แก๊ส) เข้าสู่อาคาร ตำแหน่งโดยประมาณของการวางท่อภายนอกหลักและโครงสร้าง (บ่อวาล์ว ฯลฯ) ไฟฟ้าและระบบอัจฉริยะส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อสายสาธารณูปโภค (ไฟฟ้า การสื่อสาร เคเบิลทีวี ฯลฯ) เข้าสู่อาคาร ตำแหน่งโดยประมาณของการวางสายภายนอกหลักและโครงสร้าง (บ่อตรวจสอบ ฯลฯ)",
    },
    "项目经理组织、设计主持人主持、各专业负责人参加（第三阶段开始日）：第二阶段情况总结；检查《专业配合进度表》执行情况，针对各专业配合资料进行协调，解决上阶段配合过程中的问题；安排下一阶段工作（本阶段各专业提资料如与第二阶段相比只是局部修改，可采用“局部纸质资料+完整电子文件”的方式，局部纸质资料应反映全部修改内容并与完整电子文件内容相一致）": {
        "en": "Organized by the project manager, presided over by the design director, attended by the persons in charge of all disciplines (on the starting day of the third stage): summary of the second stage; check the implementation of the Professional Coordination Schedule, coordinate the coordination materials of all disciplines, and resolve the problems in the coordination process of the previous stage; arrange the work of the next stage (if the materials submitted by each discipline in this stage are only partially modified compared with the second stage, the method of \"partial paper materials + complete electronic documents\" may be adopted; the partial paper materials shall reflect all the modifications and be consistent with the contents of the complete electronic documents)",
        "ja": "プロジェクトマネージャーが組織、設計主催者が主宰、各専門責任者が参加（第三段階開始日）：第二段階の状況総括；《専門配合進度表》の実行状況を検査、各専門の配合資料について調整し、前段階の配合過程の問題を解決；次段階の作業を手配（本段階の各専門提出資料が第二段階と比べて部分的な修正のみの場合、「部分紙資料+完全電子ファイル」方式を採用でき、部分紙資料は全部の修正内容を反映し、完全電子ファイルの内容と一致しなければならない）",
        "ko": "프로젝트 매니저가 조직하고 설계 주재자가 주재하며 각 전문 책임자가 참석(제3단계 시작일): 제2단계 상황 총괄; 《전문 협력 일정표》의 실행 상황을 점검하고, 각 전문의 협력 자료에 대해 조정하며, 이전 단계 협력 과정의 문제를 해결; 다음 단계 작업을 배치(이 단계 각 전문 제출 자료가 제2단계와 비교해 부분 수정에 그친다면 \"부분 종이 자료 + 완전 전자 파일\" 방식을 채택할 수 있으며, 부분 종이 자료는 전체 수정 내용을 반영하고 완전 전자 파일 내용과 일치해야 함)",
        "fr": "Organisée par le chef de projet, présidée par le directeur de conception, avec la participation des responsables de toutes les disciplines (le jour de début de la troisième phase) : bilan de la deuxième phase ; vérification de l'exécution du Calendrier de coordination professionnelle, coordination des documents de coordination de toutes les disciplines et résolution des problèmes survenus au cours de la phase précédente ; organisation des travaux de la phase suivante (si les documents fournis par chaque discipline à cette phase ne comportent que des modifications partielles par rapport à la deuxième phase, la méthode des \"documents papier partiels + documents électroniques complets\" peut être adoptée ; les documents papier partiels doivent refléter toutes les modifications et être cohérents avec le contenu des documents électroniques complets)",
        "de": "Organisiert durch den Projektmanager, geleitet durch den Entwurfsleiter, mit Teilnahme der Verantwortlichen aller Disziplinen (am Starttag der dritten Phase): Zusammenfassung der zweiten Phase; Prüfung der Umsetzung des Fachkoordinationsplans, Koordination der Koordinationsunterlagen aller Disziplinen und Lösung der Probleme des vorhergehenden Koordinationsprozesses; Organisation der Arbeiten der nächsten Phase (wenn die in dieser Phase von jeder Disziplin eingereichten Unterlagen im Vergleich zur zweiten Phase nur teilweise geändert wurden, kann die Methode \"teilweise Papierunterlagen + vollständige elektronische Dokumente\" angewendet werden; die teilweisen Papierunterlagen müssen alle Änderungen widerspiegeln und mit dem Inhalt der vollständigen elektronischen Dokumente übereinstimmen)",
        "es": "Organizada por el director de proyecto, presidida por el director de diseño y con la participación de los responsables de todas las especialidades (el día de inicio de la tercera fase): resumen de la segunda fase; comprobación de la ejecución del Cronograma de Coordinación Profesional, coordinación de los documentos de coordinación de todas las especialidades y resolución de los problemas surgidos en el proceso de coordinación de la fase anterior; organización de los trabajos de la fase siguiente (si los documentos aportados por cada especialidad en esta fase solo presentan modificaciones parciales respecto de la segunda fase, puede adoptarse el método de \"documentos parciales en papel + documentos electrónicos completos\"; los documentos parciales en papel deben reflejar todas las modificaciones y ser coherentes con el contenido de los documentos electrónicos completos)",
        "pt": "Organizada pelo gerente de projeto, presidida pelo diretor de projeto e com a participação dos responsáveis de todas as especialidades (no dia de início da terceira fase): resumo da segunda fase; verificação da execução do Cronograma de Coordenação Profissional, coordenação dos documentos de coordenação de todas as especialidades e resolução dos problemas surgidos no processo de coordenação da fase anterior; organização dos trabalhos da fase seguinte (se os documentos fornecidos por cada especialidade nesta fase tiverem apenas modificações parciais em relação à segunda fase, pode-se adotar o método de \"documentos parciais em papel + documentos eletrônicos completos\"; os documentos parciais em papel devem refletir todas as modificações e ser coerentes com o conteúdo dos documentos eletrônicos completos)",
        "ru": "Организуется руководителем проекта, ведётся руководителем проектирования, с участием ответственных всех специальностей (в день начала третьей стадии): итоги второй стадии; проверка выполнения Графика координации специальностей, координация материалов взаимодействия всех специальностей и решение проблем процесса взаимодействия предыдущей стадии; организация работ следующей стадии (если материалы, передаваемые каждой специальностью на этой стадии, отличаются от второй стадии лишь частичными изменениями, может применяться способ «частичные бумажные материалы + полные электронные документы»; частичные бумажные материалы должны отражать все изменения и соответствовать содержанию полных электронных документов)",
        "ar": "ينظمه مدير المشروع، ويرأسه مدير التصميم، وبمشاركة مسؤولي جميع التخصصات (في يوم بداية المرحلة الثالثة): تلخيص أوضاع المرحلة الثانية؛ والتحقق من تنفيذ جدول التنسيق المهني، وتنسيق وثائق التنسيق لجميع التخصصات، وحل المشكلات التي ظهرت في عملية التنسيق في المرحلة السابقة؛ وترتيب أعمال المرحلة التالية (إذا كانت الوثائق المقدمة من كل تخصص في هذه المرحلة مجرد تعديلات جزئية مقارنة بالمرحلة الثانية، فيمكن اعتماد أسلوب \"الوثائق الورقية الجزئية + الوثائق الإلكترونية الكاملة\"؛ ويجب أن تعكس الوثائق الورقية الجزئية جميع التعديلات وتتوافق مع محتوى الوثائق الإلكترونية الكاملة)",
        "it": "Organizzata dal project manager, presieduta dal direttore di progettazione, con la partecipazione dei responsabili di tutte le discipline (nel giorno di inizio della terza fase): sintesi della seconda fase; verifica dell'attuazione del Programma di Coordinamento Professionale, coordinamento dei documenti di coordinamento di tutte le discipline e risoluzione dei problemi emersi nel processo di coordinamento della fase precedente; organizzazione dei lavori della fase successiva (se i documenti forniti da ogni disciplina in questa fase presentano solo modifiche parziali rispetto alla seconda fase, si può adottare il metodo dei \"documenti cartacei parziali + documenti elettronici completi\"; i documenti cartacei parziali devono riflettere tutte le modifiche ed essere coerenti con il contenuto dei documenti elettronici completi)",
        "vi": "Do giám đốc dự án tổ chức, chủ trì thiết kế chủ trì, các chuyên ngành phụ trách tham dự (ngày bắt đầu giai đoạn 3): tổng kết tình hình giai đoạn 2; kiểm tra việc thực hiện Biểu tiến độ phối hợp chuyên ngành, điều phối tài liệu phối hợp của các chuyên ngành, giải quyết vấn đề trong quá trình phối hợp giai đoạn trước; bố trí công việc giai đoạn tiếp theo (nếu tài liệu các chuyên ngành nộp ở giai đoạn này chỉ sửa đổi cục bộ so với giai đoạn 2, có thể áp dụng cách \"tài liệu giấy cục bộ + hồ sơ điện tử đầy đủ\"; tài liệu giấy cục bộ phải phản ánh toàn bộ nội dung sửa đổi và thống nhất với nội dung hồ sơ điện tử đầy đủ)",
        "th": "จัดโดยผู้จัดการโครงการ ดำเนินรายการโดยหัวหน้าฝ่ายออกแบบ โดยมีผู้รับผิดชอบทุกสาขาเข้าร่วม (ในวันเริ่มระยะที่สาม): สรุปสถานการณ์ระยะที่สอง ตรวจสอบการดำเนินการตามตารางการประสานงานวิชาชีพ ประสานงานเอกสารการประสานของทุกสาขา และแก้ไขปัญหาที่เกิดขึ้นในกระบวนการประสานงานระยะก่อนหน้า จัดเตรียมงานระยะถัดไป (หากเอกสารที่แต่ละสาขาส่งในระยะนี้มีการแก้ไขเพียงบางส่วนเมื่อเทียบกับระยะที่สอง สามารถใช้วิธี \"เอกสารกระดาษบางส่วน + เอกสารอิเล็กทรอนิกส์ฉบับสมบูรณ์\" โดยเอกสารกระดาษบางส่วนต้องสะท้อนการแก้ไขทั้งหมดและสอดคล้องกับเนื้อหาของเอกสารอิเล็กทรอนิกส์ฉบับสมบูรณ์)",
    },
    "建筑专业提出（在第一版建筑作业图基础上）：全面完整反映第二阶段各专业提资料、协调配合成果；完善补充第一版作业图内容（准确的防火分区图、人防平面及口部布置与详细做法、全部房间名称、建筑专业本身的三道尺寸标注等）；对特定房间提供完整详图放大图（核心筒、客房、住宅单元放大图）；对预留竖井标注使用专业（水、暖、强电、弱电）。接受专业：总图、结构、水、暖、电、智能化、经济": {
        "en": "Proposed by the architectural discipline (on the basis of the first version of architectural working drawings): fully and completely reflect the materials submitted by all disciplines and the coordination results of the second stage; improve and supplement the contents of the first version of working drawings (accurate fire compartmentation drawings, civil air defense plans with entrance layouts and detailed practices, all room names, the architectural discipline's own three-tier dimensioning, etc.); provide complete enlarged detail drawings for specific rooms (core tube, guest rooms, residential unit enlarged drawings); annotate the using disciplines for reserved shafts (plumbing, HVAC, strong current, weak current). Receiving disciplines: site planning, structural, plumbing, HVAC, electrical, intelligent systems, economics",
        "ja": "建築専門が提出（第一版建築作業図に基づく）：第二段階の各専門提出資料と調整配合成果を全面的に反映；第一版作業図の内容を改善・補充（正確な防火区画図、人防平面及び口部配置と詳細做法、全室名、建築専門自身の三道寸法表示など）；特定の部屋には完全な詳細拡大図（コアチューブ、客室、住宅ユニット拡大図）を提供；予備シャフトに使用専門（水、暖、強電、弱電）を注記。受領専門：総図、構造、水、暖、電、智能化、経済",
        "ko": "건축 전문이 제출(제1판 건축 작업도 기반): 제2단계 각 전문 제출 자료와 조정·협력 성과를 전면적으로 반영; 제1판 작업도 내용을 보완·개선(정확한 방화 구획도, 인방 평면 및 출입구 배치와 상세 공법, 전체 실명, 건축 전문 자체의 3중 치수 표기 등); 특정 실에는 완전한 상세 확대도(코어 튜브, 객실, 주거 유닛 확대도) 제공; 예비 샤프트에 사용 전문(급배수, 난방, 강전, 약전)을 표기. 수령 전문: 총도, 구조, 급배수, 난방, 전기, 지능화, 경제",
        "fr": "Présenté par la discipline architecture (sur la base de la première version des plans de travail d'architecture) : refléter de manière complète les documents fournis par toutes les disciplines et les résultats de coordination de la deuxième phase ; améliorer et compléter le contenu de la première version des plans de travail (plans précis de compartimentage coupe-feu, plans de défense civile avec dispositions des entrées et pratiques détaillées, toutes les appellations de pièces, le triple cotage propre à la discipline architecture, etc.) ; fournir des plans agrandis détaillés complets pour des pièces spécifiques (noyau central, chambres, plans agrandis d'unités résidentielles) ; indiquer les disciplines utilisatrices des gaines réservées (plomberie, CVC, courants forts, courants faibles). Disciplines réceptrices : aménagement du site, structure, plomberie, CVC, électricité, systèmes intelligents, économie",
        "de": "Vorgelegt von der Architekturdisziplin (auf der Grundlage der ersten Fassung der Architektur-Werkspläne): die in der zweiten Phase von allen Disziplinen eingereichten Unterlagen und Koordinationsergebnisse vollständig widerspiegeln; die Inhalte der ersten Fassung der Werkspläne verbessern und ergänzen (genaue Brandabschnittspläne, Schutzraumbereiche mit Eingangsanordnungen und detaillierten Ausführungen, alle Raumbezeichnungen, die dreistufige Bemaßung der Architekturdisziplin selbst usw.); für bestimmte Räume vollständige Detailvergrößerungen liefern (Kern, Gästezimmer, Wohnungsgrundrissvergrößerungen); für zurückbehaltene Schächte die nutzenden Disziplinen kennzeichnen (Sanitär, Heizung-Lüftung, Starkstrom, Schwachstrom). Empfangende Disziplinen: Standortplanung, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de arquitectura (sobre la base de la primera versión de los planos de trabajo de arquitectura): reflejar de forma completa los documentos aportados por todas las especialidades y los resultados de coordinación de la segunda fase; mejorar y complementar el contenido de la primera versión de los planos de trabajo (planos precisos de compartimentación contra incendios, planos de defensa civil con disposición de accesos y soluciones detalladas, todas las denominaciones de locales, el triple acotado propio de la especialidad de arquitectura, etc.); facilitar planos de detalle ampliados completos para locales específicos (núcleo central, habitaciones, planos ampliados de unidades residenciales); indicar las especialidades usuarias de los patinillos reservados (fontanería, climatización, corrientes fuertes, corrientes débiles). Especialidades receptoras: urbanización, estructuras, fontanería, climatización, electricidad, sistemas inteligentes, economía",
        "pt": "Apresentado pela especialidade de arquitetura (com base na primeira versão dos desenhos de trabalho de arquitetura): refletir de forma completa os documentos fornecidos por todas as especialidades e os resultados de coordenação da segunda fase; aprimorar e complementar o conteúdo da primeira versão dos desenhos de trabalho (plantas precisas de compartimentação corta-fogo, plantas de defesa civil com disposições de acessos e práticas detalhadas, todas as denominações de ambientes, o triple dimensionamento próprio da especialidade de arquitetura, etc.); fornecer desenhos de detalhe ampliados completos para ambientes específicos (núcleo central, quartos, desenhos ampliados de unidades residenciais); indicar as especialidades usuárias dos shafts reservados (hidráulica, climatização, corrente forte, corrente fraca). Especialidades receptoras: urbanismo, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes, economia",
        "ru": "Представляется архитектурной специальностью (на основе первой редакции архитектурных рабочих чертежей): полностью отразить материалы, переданные всеми специальностями, и результаты координации второй стадии; дополнить и уточнить содержание первой редакции рабочих чертежей (точные схемы противопожарных отсеков, планы гражданской обороны с расположением входов и деталями исполнения, все наименования помещений, собственная трёхрядная простановка размеров архитектурной специальности и т.п.); для отдельных помещений представить полные увеличенные детали (ядро, гостиничные номера, увеличенные планы жилых ячеек); для резервируемых шахт указать использующие специальности (водоснабжение, отопление и вентиляция, сильные токи, слабые токи). Получающие специальности: генплан, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы, экономика",
        "ar": "يقدمه تخصص العمارة (على أساس النسخة الأولى من المخططات العملية المعمارية): عكس وثائق جميع التخصصات ونتائج التنسيق في المرحلة الثانية بشكل كامل وشامل؛ وتحسين واستكمال محتوى النسخة الأولى من المخططات العملية (مخططات تقسيمات الحماية من الحرائق الدقيقة، ومخططات الدفاع المدني مع توزيع المداخل والطرق التفصيلية، وجميع أسماء الغرف، والتحديد الثلاثي الأبعاد الخاص بتخصص العمارة وغيرها)؛ وتقديم مخططات تفصيلية موسعة كاملة لغرف محددة (النواة، الغرف الفندقية، مخططات موسعة للوحدات السكنية)؛ وتحديد التخصصات المستخدمة للأعمدة المحجوزة (السباكة، التدفئة والتهوية، التيار القوي، التيار الضعيف). التخصصات المستلمة: تخطيط المواقع والإنشاءات والسباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية والاقتصاد",
        "it": "Presentato dalla disciplina architettonica (sulla base della prima versione degli elaborati di lavoro architettonici): riflettere in modo completo e integrale i documenti forniti da tutte le discipline e i risultati di coordinamento della seconda fase; migliorare e integrare i contenuti della prima versione degli elaborati di lavoro (precisi elaborati di compartimentazione antincendio, piani di difesa civile con disposizioni degli ingressi e pratiche dettagliate, tutte le denominazioni dei locali, la quotatura tripla propria della disciplina architettonica, ecc.); fornire elaborati di dettaglio ingranditi completi per locali specifici (nucleo centrale, camere, elaborati ingranditi di unità residenziali); annotare le discipline utilizzatrici dei cavedi riservati (idrici, HVAC, corrente forte, corrente debole). Discipline riceventi: pianificazione del sito, strutture, idrici, HVAC, elettrici, sistemi intelligenti, economia",
        "vi": "Chuyên ngành kiến trúc đưa ra (trên cơ sở bản vẽ làm việc kiến trúc phiên bản 1): phản ánh đầy đủ, toàn diện tài liệu các chuyên ngành nộp ở giai đoạn 2 và thành quả phối hợp, điều phối; hoàn thiện, bổ sung nội dung bản vẽ làm việc phiên bản 1 (bản vẽ phân khu chống cháy chính xác, mặt bằng phòng hộ dân sự với bố trí cửa mở và cấu tạo chi tiết, toàn bộ tên phòng, ghi kích thước ba lớp của riêng chuyên ngành kiến trúc, v.v.); cung cấp bản vẽ phóng to chi tiết đầy đủ cho phòng đặc biệt (lõi thang máy, phòng khách sạn, bản vẽ phóng to đơn vị nhà ở); ghi chú chuyên ngành sử dụng cho hộp kỹ thuật chờ (cấp thoát nước, điều hòa, điện mạnh, điện yếu). Chuyên ngành nhận: tổng mặt bằng, kết cấu, cấp thoát nước, điều hòa, điện, hệ thống thông minh, kinh tế",
        "th": "สาขาสถาปัตยกรรมนำเสนอ (บนพื้นฐานแบบปฏิบัติงานสถาปัตยกรรมฉบับแรก): สะท้อนเอกสารที่ทุกสาขาส่งและผลการประสานงานระยะที่สองอย่างครบถ้วนสมบูรณ์ ปรับปรุงและเพิ่มเติมเนื้อหาของแบบปฏิบัติงานฉบับแรก (แบบแบ่งเขตอัคคีภัยที่แม่นยำ แบบป้องกันภัยทางอากาศพร้อมการจัดวางทางเข้าและวิธีการโดยละเอียด ชื่อห้องทั้งหมด การกำหนดขนาดสามชั้นของสาขาสถาปัตยกรรมเอง ฯลฯ) จัดทำแบบขยายรายละเอียดที่สมบูรณ์สำหรับห้องเฉพาะ (แกนอาคาร ห้องพัก แบบขยายหน่วยที่อยู่อาศัย) ระบุสาขาที่ใช้ช่องเดินท่อสำรอง (ประปา ระบบปรับอากาศ ไฟฟ้าแรง ไฟฟ้าอ่อน) สาขาที่รับ: วางผังบริเวณ โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ เศรษฐศาสตร์",
    },
    "总图专业提出（在第二阶段基础上）：全面完整反映第二阶段各专业提资料、协调配合成果；完善补充第一版作业图内容；标注出所有红线区域内的已有和新建建筑、构筑物、道路、绿化等的名称、布置与定位，以及市政管线（管沟）的平面位置尺寸。接受专业：建筑、结构、水、暖、电、智能化、经济": {
        "en": "Proposed by the site planning discipline (on the basis of the second stage): fully and completely reflect the materials submitted by all disciplines and the coordination results of the second stage; improve and supplement the contents of the first version of working drawings; mark the names, layout and positioning of all existing and newly built buildings, structures, roads and greenery within the red line area, as well as the plan positions and dimensions of municipal pipelines (trenches). Receiving disciplines: architecture, structural, plumbing, HVAC, electrical, intelligent systems, economics",
        "ja": "総図専門が提出（第二段階に基づく）：第二段階の各専門提出資料と調整配合成果を全面的に反映；第一版作業図の内容を改善・補充；境界線内の既存及び新築建築、構築物、道路、緑化などの名称、配置と位置、および市政配管（管路溝）の平面位置・寸法をすべて注記。受領専門：建築、構造、水、暖、電、智能化、経済",
        "ko": "총도 전문이 제출(제2단계 기반): 제2단계 각 전문 제출 자료와 조정·협력 성과를 전면적으로 반영; 제1판 작업도 내용을 보완·개선; 경계선 내 모든 기존 및 신축 건축물, 구조물, 도로, 조경 등의 명칭, 배치와 위치, 그리고 시정 배관(관로구)의 평면 위치·치수를 모두 표기. 수령 전문: 건축, 구조, 급배수, 난방, 전기, 지능화, 경제",
        "fr": "Présenté par la discipline aménagement du site (sur la base de la deuxième phase) : refléter de manière complète et intégrale les documents fournis par toutes les disciplines et les résultats de coordination de la deuxième phase ; améliorer et compléter le contenu de la première version des plans de travail ; indiquer les noms, dispositions et implantations de tous les bâtiments, ouvrages, voiries et espaces verts existants et nouveaux dans la zone du périmètre, ainsi que les positions et dimensions en plan des réseaux municipaux (tranchées). Disciplines réceptrices : architecture, structure, plomberie, CVC, électricité, systèmes intelligents, économie",
        "de": "Vorgelegt von der Standortplanungsdisziplin (auf der Grundlage der zweiten Phase): die in der zweiten Phase von allen Disziplinen eingereichten Unterlagen und Koordinationsergebnisse vollständig widerspiegeln; die Inhalte der ersten Fassung der Werkspläne verbessern und ergänzen; alle bestehenden und neuen Gebäude, Bauwerke, Straßen und Grünanlagen innerhalb des Rotlinienbereichs mit Namen, Anordnung und Position sowie die Planpositionen und -abmessungen der städtischen Leitungen (Kanäle) kennzeichnen. Empfangende Disziplinen: Architektur, Tragwerk, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de urbanización (sobre la base de la segunda fase): reflejar de forma completa los documentos aportados por todas las especialidades y los resultados de coordinación de la segunda fase; mejorar y complementar el contenido de la primera versión de los planos de trabajo; indicar los nombres, disposición y ubicación de todos los edificios, obras, viales y zonas verdes existentes y de nueva construcción dentro de la zona delimitada, así como las posiciones y dimensiones en planta de las redes municipales (zanjas). Especialidades receptoras: arquitectura, estructuras, fontanería, climatización, electricidad, sistemas inteligentes, economía",
        "pt": "Apresentado pela especialidade de urbanismo (com base na segunda fase): refletir de forma completa e integral os documentos fornecidos por todas as especialidades e os resultados de coordenação da segunda fase; aprimorar e complementar o conteúdo da primeira versão dos desenhos de trabalho; indicar os nomes, disposições e posicionamentos de todos os edifícios, obras, vias e áreas verdes existentes e novos dentro da área delimitada, bem como as posições e dimensões em planta das redes municipais (valas). Especialidades receptoras: arquitetura, estruturas, hidráulica, climatização, elétrica, sistemas inteligentes, economia",
        "ru": "Представляется специальностью генплана (на основе второй стадии): полностью отразить материалы, переданные всеми специальностями, и результаты координации второй стадии; дополнить и уточнить содержание первой редакции рабочих чертежей; указать наименования, расположение и привязку всех существующих и новых зданий, сооружений, дорог и озеленения в пределах красной линии, а также положение и размеры в плане городских сетей (каналов). Получающие специальности: архитектура, конструкции, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы, экономика",
        "ar": "يقدمه تخصص تخطيط المواقع (على أساس المرحلة الثانية): عكس وثائق جميع التخصصات ونتائج التنسيق في المرحلة الثانية بشكل كامل وشامل؛ وتحسين واستكمال محتوى النسخة الأولى من المخططات العملية؛ وتحديد أسماء وتوزيع ومواقع جميع المباني والمنشآت والطرق والمساحات الخضراء القائمة والجديدة داخل منطقة خط الحدود، بالإضافة إلى مواقع وأبعاد شبكات المرافق البلدية (الخنادق) في المخطط. التخصصات المستلمة: العمارة والإنشاءات والسباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية والاقتصاد",
        "it": "Presentato dalla disciplina di pianificazione del sito (sulla base della seconda fase): riflettere in modo completo e integrale i documenti forniti da tutte le discipline e i risultati di coordinamento della seconda fase; migliorare e integrare i contenuti della prima versione degli elaborati di lavoro; indicare nomi, disposizioni e posizionamenti di tutti gli edifici, opere, strade e aree verdi esistenti e nuovi nell'area del limite, nonché le posizioni e le dimensioni planimetriche delle reti comunali (canali). Discipline riceventi: architettura, strutture, idrici, HVAC, elettrici, sistemi intelligenti, economia",
        "vi": "Chuyên ngành tổng mặt bằng đưa ra (trên cơ sở giai đoạn 2): phản ánh đầy đủ, toàn diện tài liệu các chuyên ngành nộp và thành quả phối hợp, điều phối ở giai đoạn 2; hoàn thiện, bổ sung nội dung bản vẽ làm việc phiên bản 1; ghi chú tên, bố trí và định vị toàn bộ công trình, kết cấu, đường, cây xanh hiện hữu và mới xây trong khu vực ranh giới, cùng vị trí và kích thước mặt bằng đường ống hạ tầng (rãnh kỹ thuật). Chuyên ngành nhận: kiến trúc, kết cấu, cấp thoát nước, điều hòa, điện, hệ thống thông minh, kinh tế",
        "th": "สาขาวางผังบริเวณนำเสนอ (บนพื้นฐานระยะที่สอง): สะท้อนเอกสารที่ทุกสาขาส่งและผลการประสานงานระยะที่สองอย่างครบถ้วนสมบูรณ์ ปรับปรุงและเพิ่มเติมเนื้อหาของแบบปฏิบัติงานฉบับแรก ระบุชื่อ การจัดวาง และตำแหน่งของอาคาร โครงสร้าง ถนน และพื้นที่สีเขียวที่มีอยู่และสร้างใหม่ทั้งหมดภายในพื้นที่เขต รวมถึงตำแหน่งและขนาดผังของท่อสาธารณูปโภค (ร่องท่อ) สาขาที่รับ: สถาปัตยกรรม โครงสร้าง ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ เศรษฐศาสตร์",
    },
    "结构专业提出：根据第二阶段中各专业反提资料和建筑第二版作业图，提出补充和调整后的楼面、屋面结构平面图（模板图）；梁、板、柱、墙的截面与尺寸（包括定位尺寸）图；特殊结构的构造做法和详图。接受专业：建筑、总图、水、暖、电、智能化、经济": {
        "en": "Proposed by the structural discipline: according to the feedback materials of all disciplines and the second version of architectural working drawings in the second stage, present the supplemented and adjusted structural plans (formwork drawings) of floors and roofs; drawings of the sections and dimensions (including positioning dimensions) of beams, slabs, columns and walls; construction practices and details of special structures. Receiving disciplines: architecture, site planning, plumbing, HVAC, electrical, intelligent systems, economics",
        "ja": "構造専門が提出：第二段階の各専門反提資料と建築第二版作業図に基づき、補充・調整後の床・屋根構造平面図（型枠図）を提出；梁、板、柱、壁の断面と寸法（位置寸法を含む）図；特殊構造の工法と詳細図。受領専門：建築、総図、水、暖、電、智能化、経済",
        "ko": "구조 전문이 제출: 제2단계 각 전문 반제 자료와 건축 제2판 작업도에 따라, 보완·조정된 바닥·지붕 구조 평면도(거푸집 도면)를 제출; 보, 슬래브, 기둥, 벽의 단면과 치수(위치 치수 포함) 도면; 특수 구조의 시공법과 상세도. 수령 전문: 건축, 총도, 급배수, 난방, 전기, 지능화, 경제",
        "fr": "Présenté par la discipline structure : sur la base des documents de retour de toutes les disciplines et de la deuxième version des plans de travail d'architecture de la deuxième phase, présenter les plans de structure complétés et ajustés des planchers et des toitures (plans de coffrage) ; les plans des sections et dimensions (y compris les cotes de positionnement) des poutres, dalles, poteaux et murs ; les pratiques constructives et les détails des structures spéciales. Disciplines réceptrices : architecture, aménagement du site, plomberie, CVC, électricité, systèmes intelligents, économie",
        "de": "Vorgelegt von der Tragwerksdisziplin: auf der Grundlage der Rückmeldungsunterlagen aller Disziplinen und der zweiten Fassung der Architektur-Werkspläne in der zweiten Phase die ergänzten und angepassten Strukturpläne (Schalpläne) der Decken und Dächer vorlegen; Zeichnungen der Querschnitte und Abmessungen (einschließlich Positionsmaße) von Balken, Platten, Stützen und Wänden; Konstruktionsweisen und Details von Sonderkonstruktionen. Empfangende Disziplinen: Architektur, Standortplanung, Sanitär, Heizung-Lüftung, Elektro, Gebäudeautomation, Wirtschaftlichkeit",
        "es": "Presentado por la especialidad de estructuras: según los documentos de respuesta de todas las especialidades y la segunda versión de los planos de trabajo de arquitectura de la segunda fase, presentar los planos de estructura de forjados y cubiertas completados y ajustados (planos de encofrado); los planos de las secciones y dimensiones (incluidas las cotas de posición) de vigas, losas, pilares y muros; las soluciones constructivas y los detalles de las estructuras especiales. Especialidades receptoras: arquitectura, urbanización, fontanería, climatización, electricidad, sistemas inteligentes, economía",
        "pt": "Apresentado pela especialidade de estruturas: de acordo com os documentos de retorno de todas as especialidades e a segunda versão dos desenhos de trabalho de arquitetura da segunda fase, apresentar as plantas estruturais de pisos e coberturas complementadas e ajustadas (plantas de formas); os desenhos das seções e dimensões (incluindo cotas de posicionamento) de vigas, lajes, pilares e paredes; as práticas construtivas e os detalhes das estruturas especiais. Especialidades receptoras: arquitetura, urbanismo, hidráulica, climatização, elétrica, sistemas inteligentes, economia",
        "ru": "Представляется конструкторской специальностью: на основе ответных материалов всех специальностей и второй редакции архитектурных рабочих чертежей второй стадии представить дополненные и скорректированные планы конструкций перекрытий и покрытий (опалубочные чертежи); чертежи сечений и размеров (включая привязочные размеры) балок, плит, колонн и стен; конструктивные решения и детали особых конструкций. Получающие специальности: архитектура, генплан, водоснабжение, отопление и вентиляция, электроснабжение, интеллектуальные системы, экономика",
        "ar": "يقدمه تخصص الإنشاءات: بناءً على وثائق الرد لجميع التخصصات والنسخة الثانية من المخططات العملية المعمارية في المرحلة الثانية، تقديم مخططات الإنشاءات المعدلة والمستكملة للأرضيات والأسقف (مخططات القوالب)؛ ومخططات المقاطع والأبعاد (بما في ذلك أبعاد التحديد) للكمرات والألواح والأعمدة والجدران؛ وطرق الإنشاء والتفاصيل للإنشاءات الخاصة. التخصصات المستلمة: العمارة وتخطيط المواقع والسباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية والاقتصاد",
        "it": "Presentato dalla disciplina strutturale: in base ai documenti di riscontro di tutte le discipline e alla seconda versione degli elaborati di lavoro architettonici della seconda fase, presentare le piante strutturali integrate e adeguate di solai e coperture (piante di cassero); gli elaborati delle sezioni e dimensioni (incluse le quote di posizionamento) di travi, solai, pilastri e pareti; le pratiche costruttive e i dettagli delle strutture speciali. Discipline riceventi: architettura, pianificazione del sito, idrici, HVAC, elettrici, sistemi intelligenti, economia",
        "vi": "Chuyên ngành kết cấu đưa ra: căn cứ tài liệu phản hồi của các chuyên ngành và bản vẽ làm việc kiến trúc phiên bản 2 ở giai đoạn 2, đưa ra bản vẽ mặt bằng kết cấu (bản vẽ cốp pha) sàn và mái được bổ sung, điều chỉnh; bản vẽ mặt cắt và kích thước (kể cả kích thước định vị) của dầm, sàn, cột, tường; cấu tạo và chi tiết của kết cấu đặc biệt. Chuyên ngành nhận: kiến trúc, tổng mặt bằng, cấp thoát nước, điều hòa, điện, hệ thống thông minh, kinh tế",
        "th": "สาขาโครงสร้างนำเสนอ: ตามเอกสารตอบกลับของทุกสาขาและแบบปฏิบัติงานสถาปัตยกรรมฉบับที่สองในระยะที่สอง นำเสนอแบบผังโครงสร้างของพื้นและหลังคาที่เสริมและปรับแล้ว (แบบแบบหล่อ) แบบหน้าตัดและขนาด (รวมถึงขนาดกำหนดตำแหน่ง) ของคาน พื้น เสา และผนัง วิธีการก่อสร้างและรายละเอียดของโครงสร้างพิเศษ สาขาที่รับ: สถาปัตยกรรม วางผังบริเวณ ประปา ระบบปรับอากาศ ไฟฟ้า ระบบอัจฉริยะ เศรษฐศาสตร์",
    },
    "水提建筑：补齐全部管道井尺寸和位置要求；水提电：消火栓、水流指示器、报警阀、控制阀等的位置及设备和附件的控制要求；暖提建筑：补齐全部管道井尺寸和位置要求、补齐全部室外立面百页窗（新风口、排风口等）位置和尺寸；暖提电：排烟阀、加压送风阀、需要电气控制的防火阀以及自动控制阀的位置及设备和附件的控制要求；电提建筑：补齐全部管道井尺寸和位置要求；水、暖、电、智能化提结构：最大长边尺寸或直径≥1000mm的承重墙、楼板及梁上预留洞的准确位置": {
        "en": "Plumbing to architectural: complete the size and position requirements of all pipe shafts; plumbing to electrical: the positions of fire hydrants, water flow indicators, alarm valves, control valves, etc., and the control requirements of equipment and accessories; HVAC to architectural: complete the size and position requirements of all pipe shafts, and complete the positions and dimensions of all exterior louver windows (fresh air outlets, exhaust outlets, etc.); HVAC to electrical: the positions of smoke exhaust valves, pressurization air supply valves, fire dampers requiring electrical control and automatic control valves, and the control requirements of equipment and accessories; electrical to architectural: complete the size and position requirements of all pipe shafts; plumbing, HVAC, electrical and intelligent systems to structural: the accurate positions of reserved openings with maximum long-side dimension or diameter ≥1000mm in load-bearing walls, floor slabs and beams",
        "ja": "水提建築：すべてのパイプシャフト寸法と位置要求を補完；水提電：消火栓、水流指示器、警報弁、制御弁などの位置並びに設備と付属品の制御要求；暖提建築：すべてのパイプシャフト寸法と位置要求、すべての室外立面ルーバー窓（新風口、排風口など）の位置と寸法を補完；暖提電：排煙弁、加圧送風弁、電気制御が必要な防火弁および自動制御弁の位置並びに設備と付属品の制御要求；電提建築：すべてのパイプシャフト寸法と位置要求を補完；水・暖・電・智能化提構造：最大長辺寸法または直径≥1000mmの耐力壁、床版及び梁上の予備開口の正確な位置",
        "ko": "급배수가 건축에: 모든 파이프 샤프트 치수와 위치 요구를 보완; 급배수가 전기에: 소화전, 수류 지시기, 경보밸브, 제어밸브 등의 위치 및 설비와 부속품의 제어 요구; 난방이 건축에: 모든 파이프 샤프트 치수와 위치 요구, 모든 외부 입면 루버 창(신풍구, 배풍구 등)의 위치와 치수를 보완; 난방이 전기에: 배연밸브, 가압 송풍밸브, 전기 제어가 필요한 방화밸브 및 자동 제어밸브의 위치 및 설비와 부속품의 제어 요구; 전기가 건축에: 모든 파이프 샤프트 치수와 위치 요구를 보완; 급배수·난방·전기·지능화가 구조에: 최대 장변 치수 또는 직경 ≥1000mm의 내력벽, 슬래브 및 보 위 예비 구멍의 정확한 위치",
        "fr": "Plomberie à architecture : compléter les exigences de dimensions et de positions de toutes les gaines techniques ; plomberie à électricité : les positions des poteaux d'incendie, indicateurs de débit, vannes d'alarme, vannes de régulation, etc., et les exigences de commande des équipements et accessoires ; CVC à architecture : compléter les exigences de dimensions et de positions de toutes les gaines techniques et compléter les positions et dimensions de toutes les fenêtres à persiennes en façade (bouches d'air neuf, bouches d'extraction, etc.) ; CVC à électricité : les positions des clapets de désenfumage, des vannes de soufflage en surpression, des clapets coupe-feu à commande électrique et des vannes de régulation automatique, et les exigences de commande des équipements et accessoires ; électricité à architecture : compléter les exigences de dimensions et de positions de toutes les gaines techniques ; plomberie, CVC, électricité et systèmes intelligents à structure : les positions précises des réservations dont le plus grand côté ou diamètre est ≥ 1000 mm dans les murs porteurs, les dalles et les poutres",
        "de": "Sanitär an Architektur: die Größen- und Positionsanforderungen aller Installationsschächte vervollständigen; Sanitär an Elektro: die Positionen von Löschhydranten, Strömungswächtern, Alarmventilen, Regelventilen usw. und die Steuerungsanforderungen von Anlagen und Zubehör; Heizung-Lüftung an Architektur: die Größen- und Positionsanforderungen aller Installationsschächte sowie die Positionen und Abmessungen aller außenliegenden Lamellenfenster (Frischluftöffnungen, Abluftöffnungen usw.) vervollständigen; Heizung-Lüftung an Elektro: die Positionen von Rauchabzugsklappen, Druckbelüftungsklappen, elektrisch gesteuerten Brandschutzklappen und automatischen Regelventilen sowie die Steuerungsanforderungen von Anlagen und Zubehör; Elektro an Architektur: die Größen- und Positionsanforderungen aller Installationsschächte vervollständigen; Sanitär, Heizung-Lüftung, Elektro und Gebäudeautomation an Tragwerk: die genauen Positionen von Durchbrüchen mit maximaler Längsseite oder Durchmesser ≥ 1000 mm in tragenden Wänden, Geschossdecken und Balken",
        "es": "Fontanería a arquitectura: completar los requisitos de dimensiones y posiciones de todos los patinillos; fontanería a electricidad: las posiciones de las bocas de incendio, indicadores de flujo, válvulas de alarma, válvulas de control, etc., y los requisitos de control de los equipos y accesorios; climatización a arquitectura: completar los requisitos de dimensiones y posiciones de todos los patinillos y completar las posiciones y dimensiones de todas las ventanas de lamas de fachada (tomas de aire nuevo, salidas de extracción, etc.); climatización a electricidad: las posiciones de las compuertas de evacuación de humos, compuertas de impulsión de aire a presión, compuertas cortafuego de accionamiento eléctrico y válvulas de control automático, y los requisitos de control de los equipos y accesorios; electricidad a arquitectura: completar los requisitos de dimensiones y posiciones de todos los patinillos; fontanería, climatización, electricidad y sistemas inteligentes a estructuras: las posiciones precisas de los huecos cuyo lado mayor o diámetro sea ≥ 1000 mm en muros portantes, forjados y vigas",
        "pt": "Hidráulica para arquitetura: completar os requisitos de dimensões e posições de todos os shafts; hidráulica para elétrica: as posições dos hidrantes, indicadores de fluxo, válvulas de alarme, válvulas de controle, etc., e os requisitos de controle dos equipamentos e acessórios; climatização para arquitetura: completar os requisitos de dimensões e posições de todos os shafts e completar as posições e dimensões de todas as janelas de lâminas na fachada (entradas de ar novo, saídas de exaustão, etc.); climatização para elétrica: as posições das comportas de exaustão de fumaça, comportas de insuflação de ar pressurizado, comportas corta-fogo de acionamento elétrico e válvulas de controle automático, e os requisitos de controle dos equipamentos e acessórios; elétrica para arquitetura: completar os requisitos de dimensões e posições de todos os shafts; hidráulica, climatização, elétrica e sistemas inteligentes para estruturas: as posições precisas das aberturas com maior lado ou diâmetro ≥ 1000 mm em paredes estruturais, lajes e vigas",
        "ru": "Водоснабжение — архитектуре: дополнить требования к размерам и расположению всех инженерных шахт; водоснабжение — электроснабжению: расположение пожарных кранов, сигнализаторов потока, сигнальных клапанов, регулирующих клапанов и т.п., а также требования к управлению оборудованием и арматурой; отопление и вентиляция — архитектуре: дополнить требования к размерам и расположению всех инженерных шахт и дополнить расположение и размеры всех фасадных жалюзийных решёток (воздухозаборов, вытяжных отверстий и т.п.); отопление и вентиляция — электроснабжению: расположение противопожарных клапанов дымоудаления, клапанов подпора воздуха, огнезадерживающих клапанов с электроприводом и автоматических регулирующих клапанов, а также требования к управлению оборудованием и арматурой; электроснабжение — архитектуре: дополнить требования к размерам и расположению всех инженерных шахт; водоснабжение, отопление и вентиляция, электроснабжение и интеллектуальные системы — конструкциям: точное расположение отверстий с наибольшим размером стороны или диаметром ≥ 1000 мм в несущих стенах, плитах перекрытий и балках",
        "ar": "السباكة إلى العمارة: استكمال متطلبات الأبعاد والمواقع لجميع الأعمدة الفنية؛ والسباكة إلى الكهرباء: مواقع صنابير الإطفاء ومؤشرات التدفق وصمامات الإنذار وصمامات التحكم وغيرها، ومتطلبات التحكم في المعدات والملحقات؛ والتدفئة والتهوية إلى العمارة: استكمال متطلبات الأبعاد والمواقع لجميع الأعمدة الفنية، واستكمال مواقع وأبعاد جميع نوافذ الشرائح الخارجية (مداخل الهواء الجديد، مخارج الهواء العادم وغيرها)؛ والتدفئة والتهوية إلى الكهرباء: مواقع صمامات إخراج الدخان وصمامات تزويد الهواء المضغوط وصمامات الحريق التي تتطلب تحكماً كهربائياً وصمامات التحكم الآلي، ومتطلبات التحكم في المعدات والملحقات؛ والكهرباء إلى العمارة: استكمال متطلبات الأبعاد والمواقع لجميع الأعمدة الفنية؛ والسباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية إلى الإنشاءات: المواقع الدقيقة للفتحات التي يبلغ أكبر بعد أو قطر لها ≥ 1000 مم في الجدران الحاملة والألواح والكمرات",
        "it": "Idrici a architettura: completare i requisiti di dimensioni e posizioni di tutti i cavedi; idrici a elettrici: le posizioni di idranti, indicatori di flusso, valvole di allarme, valvole di controllo, ecc., e i requisiti di controllo di apparecchiature e accessori; HVAC a architettura: completare i requisiti di dimensioni e posizioni di tutti i cavedi e completare le posizioni e dimensioni di tutte le finestre a lamelle in facciata (prese d'aria nuova, bocchette di estrazione, ecc.); HVAC a elettrici: le posizioni delle serrande di evacuazione fumi, delle serrande di immissione aria in pressione, delle serrande tagliafuoco a comando elettrico e delle valvole di controllo automatico, e i requisiti di controllo di apparecchiature e accessori; elettrici a architettura: completare i requisiti di dimensioni e posizioni di tutti i cavedi; idrici, HVAC, elettrici e sistemi intelligenti a strutture: le posizioni precise delle aperture con lato maggiore o diametro ≥ 1000 mm nelle pareti portanti, nei solai e nelle travi",
        "vi": "Cấp thoát nước gửi kiến trúc: bổ sung đầy đủ yêu cầu kích thước và vị trí các hộp kỹ thuật đường ống; cấp thoát nước gửi điện: vị trí họng cứu hỏa, chỉ báo dòng chảy, van báo động, van điều khiển, v.v. và yêu cầu điều khiển thiết bị, phụ kiện; điều hòa gửi kiến trúc: bổ sung đầy đủ yêu cầu kích thước và vị trí các hộp kỹ thuật đường ống, bổ sung đầy đủ vị trí và kích thước cửa chớp mặt đứng (miệng gió tươi, miệng gió thải, v.v.); điều hòa gửi điện: vị trí van xả khói, van cấp gió tăng áp, van chống cháy cần điều khiển điện và van điều khiển tự động, cùng yêu cầu điều khiển thiết bị, phụ kiện; điện gửi kiến trúc: bổ sung đầy đủ yêu cầu kích thước và vị trí các hộp kỹ thuật đường ống; cấp thoát nước, điều hòa, điện, hệ thống thông minh gửi kết cấu: vị trí chính xác các lỗ chờ có cạnh dài lớn nhất hoặc đường kính ≥1000mm trên tường chịu lực, sàn và dầm",
        "th": "ประปาส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปาส่งให้ไฟฟ้า: ตำแหน่งหัวจ่ายน้ำดับเพลิง เครื่องแสดงการไหลของน้ำ วาล์วสัญญาณเตือน วาล์วควบคุม ฯลฯ และข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ระบบปรับอากาศส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด และเติมเต็มตำแหน่งและขนาดของช่องระบายอากาศแบบใบพัดภายนอกทั้งหมด (ช่องลมเข้า ช่องลมออก ฯลฯ) ระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งวาล์วระบายควัน วาล์วส่งลมอัดแรงดัน วาล์วกันไฟที่ต้องควบคุมด้วยไฟฟ้า และวาล์วควบคุมอัตโนมัติ รวมถึงข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ไฟฟ้าส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งที่แม่นยำของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนัก พื้น และคาน",
    },
    "水、暖、电、智能化提结构：吊挂管道和设备的楼板荷载；需要结构专业进行设计的大型设备的基础尺寸、设备重量及相关要求；水、暖专业中需要结构专业进行加固的管道固定支架的位置和推力；水、暖提建筑：排水沟的位置、尺寸、坡度等要求；水专业提建筑：特殊房间（如气体消防房间等）对建筑门、窗等洞口的要求": {
        "en": "Plumbing, HVAC, electrical and intelligent systems to structural: the floor slab loads of suspended pipes and equipment; the foundation dimensions, equipment weights and related requirements of large equipment requiring structural design; the positions and thrusts of pipe fixed supports in the plumbing and HVAC disciplines that require structural reinforcement; plumbing and HVAC to architectural: the requirements for the positions, dimensions and slopes of drainage ditches; plumbing to architectural: the requirements of special rooms (such as gas fire protection rooms, etc.) for architectural openings such as doors and windows",
        "ja": "水・暖・電・智能化提構造：吊下げ配管と設備の床版荷重；構造専門の設計が必要な大型設備の基礎寸法、設備重量及び関連要求；水・暖専門の中で構造専門による補強が必要な配管固定支持架の位置と推力；水・暖提建築：排水溝の位置、寸法、勾配などの要求；水専門提建築：特殊室（ガス消防室など）の建築ドア・窓などの開口に対する要求",
        "ko": "급배수·난방·전기·지능화가 구조에: 매달린 배관과 설비의 바닥 슬래브 하중; 구조 전문의 설계가 필요한 대형 설비의 기초 치수, 설비 중량 및 관련 요구; 급배수·난방 전문 중 구조 전문의 보강이 필요한 배관 고정 지지대의 위치와 추력; 급배수·난방이 건축에: 배수구의 위치, 치수, 기울기 등의 요구; 급배수가 건축에: 특수실(가스 소방실 등)의 건축 문·창 등 개구부에 대한 요구",
        "fr": "Plomberie, CVC, électricité et systèmes intelligents à structure : les charges sur dalles des conduits suspendus et des équipements ; les dimensions des fondations, les poids des équipements et les exigences connexes des grands équipements nécessitant une conception structurelle ; les positions et poussées des supports fixes de canalisations dans les disciplines plomberie et CVC nécessitant un renforcement structurel ; plomberie et CVC à architecture : les exigences de positions, dimensions et pentes des caniveaux d'évacuation ; plomberie à architecture : les exigences des locaux spéciaux (tels que les locaux de protection incendie par gaz, etc.) concernant les baies telles que portes et fenêtres",
        "de": "Sanitär, Heizung-Lüftung, Elektro und Gebäudeautomation an Tragwerk: die Deckenlasten hängender Rohrleitungen und Anlagen; die Gründungsabmessungen, Anlagengewichte und damit verbundenen Anforderungen großer Anlagen, die eine statische Bemessung erfordern; die Positionen und Schubkräfte von Rohrleitungsfestpunkten in den Disziplinen Sanitär und Heizung-Lüftung, die eine bauliche Verstärkung erfordern; Sanitär und Heizung-Lüftung an Architektur: die Anforderungen an Position, Abmessungen und Gefälle von Entwässerungsrinnen; Sanitär an Architektur: die Anforderungen besonderer Räume (z. B. Gaslöschanlagenräume) an Öffnungen wie Türen und Fenster",
        "es": "Fontanería, climatización, electricidad y sistemas inteligentes a estructuras: las cargas sobre forjados de los conductos colgados y equipos; las dimensiones de cimentación, pesos de los equipos y requisitos relacionados de los grandes equipos que requieren diseño estructural; las posiciones y empujes de los soportes fijos de tuberías en las especialidades de fontanería y climatización que requieren refuerzo estructural; fontanería y climatización a arquitectura: los requisitos de posiciones, dimensiones y pendientes de los canales de drenaje; fontanería a arquitectura: los requisitos de los locales especiales (como los locales de protección contra incendios por gas, etc.) sobre los huecos de puertas y ventanas",
        "pt": "Hidráulica, climatização, elétrica e sistemas inteligentes para estruturas: as cargas sobre lajes de tubulações suspensas e equipamentos; as dimensões de fundações, pesos dos equipamentos e requisitos relacionados dos grandes equipamentos que exigem projeto estrutural; as posições e empuxos dos suportes fixos de tubulações nas especialidades de hidráulica e climatização que exigem reforço estrutural; hidráulica e climatização para arquitetura: os requisitos de posições, dimensões e declividades das canaletas de drenagem; hidráulica para arquitetura: os requisitos dos ambientes especiais (como salas de proteção contra incêndio por gás, etc.) sobre vãos como portas e janelas",
        "ru": "Водоснабжение, отопление и вентиляция, электроснабжение и интеллектуальные системы — конструкциям: нагрузки на перекрытия от подвесных трубопроводов и оборудования; размеры фундаментов, массы оборудования и связанные требования для крупного оборудования, требующего расчёта конструкций; расположение и распор неподвижных опор трубопроводов в специальностях водоснабжения и отопления и вентиляции, требующих усиления конструкций; водоснабжение и отопление и вентиляция — архитектуре: требования к расположению, размерам и уклонам водоотводных лотков; водоснабжение — архитектуре: требования особых помещений (например, газового пожаротушения) к проёмам, дверям и окнам",
        "ar": "السباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية إلى الإنشاءات: الأحمال على الألواح من الأنابيب المعلقة والمعدات؛ وأبعاد الأساسات وأوزان المعدات والمتطلبات ذات الصلة للمعدات الكبيرة التي تتطلب تصميمًا إنشائياً؛ ومواقع وقوى الدفع للدعامات الثابتة للأنابيب في تخصصي السباكة والتدفئة والتهوية التي تتطلب تقوية إنشائية؛ والسباكة والتدفئة والتهوية إلى العمارة: متطلبات المواقع والأبعاد والميول لقنوات الصرف؛ والسباكة إلى العمارة: متطلبات الغرف الخاصة (مثل غرف إطفاء الحرائق بالغاز وغيرها) بشأن الفتحات مثل الأبواب والنوافذ",
        "it": "Idrici, HVAC, elettrici e sistemi intelligenti a strutture: i carichi sui solai di tubazioni sospese e apparecchiature; le dimensioni delle fondazioni, i pesi delle apparecchiature e i requisiti correlati delle grandi apparecchiature che richiedono progettazione strutturale; le posizioni e le spinte dei supporti fissi delle tubazioni nelle discipline idriche e HVAC che richiedono rinforzo strutturale; idrici e HVAC a architettura: i requisiti di posizioni, dimensioni e pendenze dei canali di drenaggio; idrici a architettura: i requisiti dei locali speciali (come le sale di protezione antincendio a gas, ecc.) sulle aperture come porte e finestre",
        "vi": "Cấp thoát nước, điều hòa, điện, hệ thống thông minh gửi kết cấu: tải trọng sàn của đường ống treo và thiết bị; kích thước móng, trọng lượng thiết bị và yêu cầu liên quan của thiết bị lớn cần kết cấu thiết kế; vị trí và lực đẩy của gối đỡ cố định đường ống trong chuyên ngành cấp thoát nước, điều hòa cần kết cấu gia cường; cấp thoát nước, điều hòa gửi kiến trúc: yêu cầu vị trí, kích thước, độ dốc của rãnh thoát nước; cấp thoát nước gửi kiến trúc: yêu cầu của phòng đặc biệt (như phòng chữa cháy khí, v.v.) đối với lỗ mở cửa, cửa sổ của kiến trúc",
        "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: น้ำหนักบรรทุกบนพื้นจากท่อแขวนและอุปกรณ์ ขนาดฐานราก น้ำหนักอุปกรณ์ และข้อกำหนดที่เกี่ยวข้องของอุปกรณ์ขนาดใหญ่ที่ต้องออกแบบโครงสร้าง ตำแหน่งและแรงดันของขายึดท่อคงที่ในสาขาประปาและระบบปรับอากาศที่ต้องเสริมโครงสร้าง ประปาและระบบปรับอากาศส่งให้สถาปัตยกรรม: ข้อกำหนดตำแหน่ง ขนาด และความลาดเอียงของรางระบายน้ำ ประปาส่งให้สถาปัตยกรรม: ข้อกำหนดของห้องพิเศษ (เช่น ห้องดับเพลิงด้วยแก๊ส ฯลฯ) ต่อช่องเปิด เช่น ประตูและหน้าต่าง",
    },
    "建筑专业提出：楼梯间、卫生间放大详图；屋面及墙身剖面或构造详图、必要的结点放大图；各个门的开启方向；非承重墙中的构造柱尺寸和位置；非承重墙过梁的尺寸和位置": {
        "en": "Proposed by the architectural discipline: enlarged details of stairwells and toilets; roof and wall sections or construction details, necessary enlarged joint details; the opening directions of each door; the dimensions and positions of structural columns in non-load-bearing walls; the dimensions and positions of lintels in non-load-bearing walls",
        "ja": "建築専門が提出：階段室、トイレの拡大詳細図；屋根及び壁体断面または構造詳細図、必要な納まり拡大図；各ドアの開き方向；非耐力壁の構造柱の寸法と位置；非耐力壁のまぐさの寸法と位置",
        "ko": "건축 전문이 제출: 계단실, 화장실 확대 상세도; 지붕 및 벽체 단면 또는 구조 상세도, 필요한 접합부 확대도; 각 문의 개폐 방향; 비내력벽의 구조 기둥 치수와 위치; 비내력벽 인방의 치수와 위치",
        "fr": "Présenté par la discipline architecture : les plans de détail agrandis des cages d'escalier et des sanitaires ; les coupes de toiture et de murs ou les détails constructifs, les plans agrandis nécessaires des nœuds ; les sens d'ouverture de chaque porte ; les dimensions et positions des poteaux de chaînage dans les murs non porteurs ; les dimensions et positions des linteaux dans les murs non porteurs",
        "de": "Vorgelegt von der Architekturdisziplin: vergrößerte Details von Treppenhäusern und Toiletten; Dach- und Wandschnitte oder Konstruktionsdetails, erforderliche vergrößerte Knotenpunkte; die Öffnungsrichtungen jeder Tür; die Abmessungen und Positionen von Aussteifungspfeilern in nicht tragenden Wänden; die Abmessungen und Positionen von Stürzen in nicht tragenden Wänden",
        "es": "Presentado por la especialidad de arquitectura: detalles ampliados de cajas de escalera y aseos; secciones de cubierta y muros o detalles constructivos, ampliaciones de nudos necesarias; los sentidos de apertura de cada puerta; las dimensiones y posiciones de los pilares de atado en muros no portantes; las dimensiones y posiciones de los dinteles en muros no portantes",
        "pt": "Apresentado pela especialidade de arquitetura: detalhes ampliados de escadas e banheiros; cortes de cobertura e paredes ou detalhes construtivos, ampliações de nós necessárias; os sentidos de abertura de cada porta; as dimensões e posições dos pilares de amarração em paredes não estruturais; as dimensões e posições dos lintéis em paredes não estruturais",
        "ru": "Представляется архитектурной специальностью: увеличенные детали лестничных клеток и санузлов; разрезы кровли и стен или конструктивные детали, необходимые укрупнённые узлы; направления открывания каждой двери; размеры и расположение конструктивных колонн в ненесущих стенах; размеры и расположение перемычек в ненесущих стенах",
        "ar": "يقدمه تخصص العمارة: تفاصيل موسعة لسلالم الحريق ودورات المياه؛ ومقاطع للسقف والجدران أو تفاصيل إنشائية، ومخططات موسعة للعُقد اللازمة؛ واتجاهات فتح كل باب؛ وأبعاد ومواقع الأعمدة الإنشائية في الجدران غير الحاملة؛ وأبعاد ومواقع العتبات في الجدران غير الحاملة",
        "it": "Presentato dalla disciplina architettonica: dettagli ingranditi di vani scala e servizi igienici; sezioni di copertura e pareti o dettagli costruttivi, necessari ingrandimenti dei nodi; i sensi di apertura di ogni porta; le dimensioni e posizioni dei pilastri di ammorsamento nelle pareti non portanti; le dimensioni e posizioni degli architravi nelle pareti non portanti",
        "vi": "Chuyên ngành kiến trúc đưa ra: bản vẽ phóng to chi tiết cầu thang, nhà vệ sinh; mặt cắt mái và tường hoặc chi tiết cấu tạo, bản vẽ phóng to mắt nút cần thiết; hướng mở của từng cửa; kích thước và vị trí cột kết cấu trong tường không chịu lực; kích thước và vị trí lanh tô trong tường không chịu lực",
        "th": "สาขาสถาปัตยกรรมนำเสนอ: แบบขยายรายละเอียดบันไดและห้องน้ำ หน้าตัดหลังคาและผนังหรือรายละเอียดก่อสร้าง แบบขยายจุดต่อที่จำเป็น ทิศทางการเปิดของประตูแต่ละบาน ขนาดและตำแหน่งเสาโครงสร้างในผนังไม่รับน้ำหนัก ขนาดและตำแหน่งทับหลังในผนังไม่รับน้ำหนัก",
    },
    "需要结构专业配合设计的水池、积水坑等详图（包括外形尺寸、厚度、预留洞、相关做法、保温要求等）": {
        "en": "Details of pools, sumps, etc. that require the structural discipline to cooperate in design (including overall dimensions, thickness, reserved openings, relevant practices, thermal insulation requirements, etc.)",
        "ja": "構造専門の協力設計が必要な水池、集水桝などの詳細図（外形寸法、厚さ、予備開口、関連做法、保温要求などを含む）",
        "ko": "구조 전문의 협력 설계가 필요한 물탱크, 집수정 등의 상세도(외형 치수, 두께, 예비 구멍, 관련 시공법, 보온 요구 등 포함)",
        "fr": "Les plans de détail des bassins, puisards, etc. nécessitant la coopération de la discipline structure dans la conception (y compris les dimensions hors tout, l'épaisseur, les réservations, les pratiques associées, les exigences d'isolation thermique, etc.)",
        "de": "Details von Becken, Sümpfen usw., die eine mitwirkende Bemessung durch die Tragwerksdisziplin erfordern (einschließlich Gesamtabmessungen, Dicke, Durchbrüchen, einschlägiger Ausführungen, Wärmedämmanforderungen usw.)",
        "es": "Los detalles de depósitos, sumideros, etc., que requieren la cooperación de la especialidad de estructuras en el diseño (incluidos los dimensiones totales, el espesor, los huecos reservados, las soluciones asociadas, los requisitos de aislamiento térmico, etc.)",
        "pt": "Os detalhes de reservatórios, poços de coleta, etc., que exigem a cooperação da especialidade de estruturas no projeto (incluindo dimensões totais, espessura, aberturas reservadas, práticas associadas, requisitos de isolamento térmico, etc.)",
        "ru": "Детали резервуаров, приямков и т.п., требующие участия конструкторской специальности в проектировании (включая габаритные размеры, толщину, отверстия, соответствующие решения, требования к теплоизоляции и т.п.)",
        "ar": "تفاصيل الأحواض وآبار التجميع وغيرها التي تتطلب تعاون تخصص الإنشاءات في التصميم (بما في ذلك الأبعاد الكلية والسماكة والفتحات المحجوزة والممارسات ذات الصلة ومتطلبات العزل الحراري وغيرها)",
        "it": "I dettagli di vasche, pozzetti, ecc., che richiedono la cooperazione della disciplina strutturale nella progettazione (inclusi dimensioni d'ingombro, spessore, aperture riservate, pratiche correlate, requisiti di isolamento termico, ecc.)",
        "vi": "Chi tiết bể nước, hố thu nước, v.v. cần chuyên ngành kết cấu phối hợp thiết kế (bao gồm kích thước bao, chiều dày, lỗ chờ, cấu tạo liên quan, yêu cầu cách nhiệt, v.v.)",
        "th": "รายละเอียดของบ่อน้ำ บ่อรวบรวมน้ำ ฯลฯ ที่ต้องให้สาขาโครงสร้างร่วมออกแบบ (รวมถึงขนาดโดยรวม ความหนา ช่องเผื่อ วิธีการที่เกี่ยวข้อง ข้อกำหนดฉนวนกันความร้อน ฯลฯ)",
    },
    "设计主持人组织、各专业负责人参加（必要时相关专业设计人参加）：第三阶段情况总结，安排下一阶段工作；检查《专业配合进度表》执行情况，针对各专业配合资料进行协调，解决上阶段配合过程中的问题": {
        "en": "Organized by the design director, attended by the persons in charge of all disciplines (designers of relevant disciplines attend when necessary): summary of the third stage and arrangement of the work of the next stage; check the implementation of the Professional Coordination Schedule, coordinate the coordination materials of all disciplines, and resolve the problems in the coordination process of the previous stage",
        "ja": "設計主催者が組織、各専門責任者が参加（必要な場合、関連専門の設計者が参加）：第三段階の状況総括、次段階の作業を手配；《専門配合進度表》の実行状況を検査、各専門の配合資料について調整し、前段階の配合過程の問題を解決",
        "ko": "설계 주재자가 조직하고 각 전문 책임자가 참석(필요시 관련 전문 설계자 참석): 제3단계 상황 총괄, 다음 단계 작업 배치; 《전문 협력 일정표》의 실행 상황 점검, 각 전문의 협력 자료에 대해 조정하며, 이전 단계 협력 과정의 문제를 해결",
        "fr": "Organisée par le directeur de conception, avec la participation des responsables de toutes les disciplines (les concepteurs des disciplines concernées participent en cas de besoin) : bilan de la troisième phase et organisation des travaux de la phase suivante ; vérification de l'exécution du Calendrier de coordination professionnelle, coordination des documents de coordination de toutes les disciplines et résolution des problèmes survenus au cours de la phase précédente",
        "de": "Organisiert durch den Entwurfsleiter, mit Teilnahme der Verantwortlichen aller Disziplinen (bei Bedarf nehmen auch Planer der betreffenden Disziplinen teil): Zusammenfassung der dritten Phase und Organisation der Arbeiten der nächsten Phase; Prüfung der Umsetzung des Fachkoordinationsplans, Koordination der Koordinationsunterlagen aller Disziplinen und Lösung der Probleme des vorhergehenden Koordinationsprozesses",
        "es": "Organizada por el director de diseño, con la participación de los responsables de todas las especialidades (los proyectistas de las especialidades implicadas participan cuando sea necesario): resumen de la tercera fase y organización de los trabajos de la fase siguiente; comprobación de la ejecución del Cronograma de Coordinación Profesional, coordinación de los documentos de coordinación de todas las especialidades y resolución de los problemas surgidos en el proceso de coordinación de la fase anterior",
        "pt": "Organizada pelo diretor de projeto, com a participação dos responsáveis de todas as especialidades (os projetistas das especialidades envolvidas participam quando necessário): resumo da terceira fase e organização dos trabalhos da fase seguinte; verificação da execução do Cronograma de Coordenação Profissional, coordenação dos documentos de coordenação de todas as especialidades e resolução dos problemas surgidos no processo de coordenação da fase anterior",
        "ru": "Организуется руководителем проектирования, с участием ответственных всех специальностей (при необходимости участвуют проектировщики соответствующих специальностей): итоги третьей стадии и организация работ следующей стадии; проверка выполнения Графика координации специальностей, координация материалов взаимодействия всех специальностей и решение проблем процесса взаимодействия предыдущей стадии",
        "ar": "ينظمه مدير التصميم، وبمشاركة مسؤولي جميع التخصصات (يشارك مصممو التخصصات ذات الصلة عند الحاجة): تلخيص أوضاع المرحلة الثالثة وترتيب أعمال المرحلة التالية؛ والتحقق من تنفيذ جدول التنسيق المهني، وتنسيق وثائق التنسيق لجميع التخصصات، وحل المشكلات التي ظهرت في عملية التنسيق في المرحلة السابقة",
        "it": "Organizzata dal direttore di progettazione, con la partecipazione dei responsabili di tutte le discipline (i progettisti delle discipline interessate partecipano quando necessario): sintesi della terza fase e organizzazione dei lavori della fase successiva; verifica dell'attuazione del Programma di Coordinamento Professionale, coordinamento dei documenti di coordinamento di tutte le discipline e risoluzione dei problemi emersi nel processo di coordinamento della fase precedente",
        "vi": "Chủ trì thiết kế tổ chức, các chuyên ngành phụ trách tham dự (nếu cần, thiết kế viên chuyên ngành liên quan tham dự): tổng kết tình hình giai đoạn 3, bố trí công việc giai đoạn tiếp theo; kiểm tra việc thực hiện Biểu tiến độ phối hợp chuyên ngành, điều phối tài liệu phối hợp của các chuyên ngành, giải quyết vấn đề trong quá trình phối hợp giai đoạn trước",
        "th": "หัวหน้าฝ่ายออกแบบจัด โดยผู้รับผิดชอบทุกสาขาเข้าร่วม (นักออกแบบสาขาที่เกี่ยวข้องเข้าร่วมเมื่อจำเป็น): สรุปสถานการณ์ระยะที่สามและจัดเตรียมงานระยะถัดไป ตรวจสอบการดำเนินการตามตารางการประสานงานวิชาชีพ ประสานงานเอกสารการประสานของทุกสาขา และแก้ไขปัญหาที่เกิดขึ้นในกระบวนการประสานงานระยะก่อนหน้า",
    },
    "设计主持人组织：对所有管道进行综合与协调，防止管道打架；验证与核对房间名称和使用功能；验证与核对防火分区；验证与核对各专业所提资料和配合过程中的协商内容是否在相关图纸中得到反映；验证各专业施工图是否符合初步设计配合原则；核对并解决管线与建筑图中的构造柱、过梁等可能存在的打架情况（本阶段 71%W～76%W 时间周期内完成）": {
        "en": "Organized by the design director: comprehensively integrate and coordinate all pipelines to prevent pipeline conflicts; verify and check room names and usage functions; verify and check fire compartmentation; verify and check whether the materials submitted by all disciplines and the consultation contents in the coordination process are reflected in the relevant drawings; verify whether the construction drawings of all disciplines conform to the preliminary design coordination principles; check and resolve possible conflicts between pipelines and structural columns, lintels, etc. in the architectural drawings (to be completed within the 71%W～76%W time period of this stage)",
        "ja": "設計主催者が組織：すべての配管を総合調整し、配管の干渉を防止；部屋名と使用機能を検証・照合；防火区画を検証・照合；各専門が提出した資料と配合過程の協議内容が関連図面に反映されているかを検証・照合；各専門の施工図が基本設計配合原則に適合するかを検証；配管と建築図中の構造柱、まぐさなどの潜在的な干渉を照合・解決（本段階 71%W～76%W の期間内に完了）",
        "ko": "설계 주재자가 조직: 모든 배관을 종합 조정하여 배관 간섭을 방지; 실명과 사용 기능을 검증·대조; 방화 구획을 검증·대조; 각 전문이 제출한 자료와 협력 과정의 협의 내용이 관련 도면에 반영되었는지 검증·대조; 각 전문의 시공도가 기본 설계 협력 원칙에 부합하는지 검증; 배관과 건축도면 중 구조 기둥, 인방 등의 잠재적 간섭을 대조·해결(이 단계 71%W~76%W 기간 내 완료)",
        "fr": "Organisée par le directeur de conception : intégrer et coordonner l'ensemble des canalisations pour éviter les conflits entre canalisations ; vérifier et contrôler les appellations des pièces et les fonctions d'usage ; vérifier et contrôler le compartimentage coupe-feu ; vérifier et contrôler si les documents fournis par toutes les disciplines et le contenu des concertations du processus de coordination sont reflétés dans les plans concernés ; vérifier si les plans d'exécution de toutes les disciplines sont conformes aux principes de coordination de l'avant-projet ; contrôler et résoudre les conflits éventuels entre les canalisations et les poteaux de chaînage, les linteaux, etc. dans les plans d'architecture (à achever dans la période 71 %W à 76 %W de cette phase)",
        "de": "Organisiert durch den Entwurfsleiter: alle Rohrleitungen integrieren und koordinieren, um Leitungsquerungen zu vermeiden; Raumbezeichnungen und Nutzungsfunktionen verifizieren und prüfen; Brandabschnitte verifizieren und prüfen; verifizieren und prüfen, ob die von allen Disziplinen eingereichten Unterlagen und die in der Koordination vereinbarten Inhalte in den betreffenden Plänen widergespiegelt sind; prüfen, ob die Ausführungspläne aller Disziplinen den Koordinationsprinzipien der Entwurfsplanung entsprechen; mögliche Kollisionen zwischen Leitungen und Aussteifungspfeilern, Stürzen usw. in den Architekturplänen prüfen und lösen (innerhalb des Zeitraums 71 %W bis 76 %W dieser Phase abzuschließen)",
        "es": "Organizada por el director de diseño: integrar y coordinar todas las canalizaciones para evitar conflictos entre ellas; verificar y comprobar las denominaciones de los locales y sus funciones de uso; verificar y comprobar la compartimentación contra incendios; verificar y comprobar si los documentos aportados por todas las especialidades y el contenido de las consultas del proceso de coordinación se reflejan en los planos correspondientes; verificar si los planos de ejecución de todas las especialidades se ajustan a los principios de coordinación del anteproyecto; comprobar y resolver los posibles conflictos entre las canalizaciones y los pilares de atado, dinteles, etc. de los planos de arquitectura (debe completarse dentro del periodo 71%W a 76%W de esta fase)",
        "pt": "Organizada pelo diretor de projeto: integrar e coordenar todas as tubulações para evitar conflitos entre elas; verificar e conferir as denominações dos ambientes e suas funções de uso; verificar e conferir a compartimentação corta-fogo; verificar e conferir se os documentos fornecidos por todas as especialidades e o conteúdo das consultas do processo de coordenação estão refletidos nos desenhos correspondentes; verificar se os desenhos de execução de todas as especialidades estão em conformidade com os princípios de coordenação do anteprojeto; conferir e resolver os possíveis conflitos entre tubulações e pilares de amarração, lintéis, etc. nos desenhos de arquitetura (a concluir dentro do período de 71%W a 76%W desta fase)",
        "ru": "Организуется руководителем проектирования: комплексно увязать все трубопроводы, исключив их пересечения; проверить и сверить наименования помещений и их функции; проверить и сверить противопожарные отсеки; проверить и сверить, отражены ли в соответствующих чертежах материалы, переданные всеми специальностями, и результаты согласований в процессе взаимодействия; проверить соответствие рабочих чертежей всех специальностей принципам координации на стадии предпроекта; сверить и устранить возможные пересечения трубопроводов с конструктивными колоннами, перемычками и т.п. на архитектурных чертежах (завершить в период 71%W–76%W данной стадии)",
        "ar": "ينظمه مدير التصميم: التنسيق الشامل لجميع خطوط الأنابيب لمنع تعارضها؛ والتحقق من أسماء الغرف ووظائف الاستخدام؛ والتحقق من تقسيمات الحماية من الحرائق؛ والتحقق مما إذا كانت الوثائق المقدمة من جميع التخصصات ومحتوى المشاورات في عملية التنسيق قد انعكست في المخططات ذات الصلة؛ والتحقق من مطابقة المخططات التنفيذية لجميع التخصصات لمبادئ تنسيق التصميم الأولي؛ ومراجعة وحل التعارضات المحتملة بين خطوط الأنابيب والأعمدة الإنشائية والعتبات وغيرها في المخططات المعمارية (يُنجز خلال الفترة الزمنية 71%W～76%W من هذه المرحلة)",
        "it": "Organizzata dal direttore di progettazione: integrare e coordinare tutte le tubazioni per evitare conflitti tra loro; verificare e controllare le denominazioni dei locali e le funzioni d'uso; verificare e controllare la compartimentazione antincendio; verificare e controllare se i documenti forniti da tutte le discipline e i contenuti delle consultazioni del processo di coordinamento sono riflessi negli elaborati pertinenti; verificare se gli elaborati esecutivi di tutte le discipline sono conformi ai principi di coordinamento del progetto preliminare; controllare e risolvere i possibili conflitti tra tubazioni e pilastri di ammorsamento, architravi, ecc. negli elaborati architettonici (da completare entro il periodo 71%W～76%W di questa fase)",
        "vi": "Chủ trì thiết kế tổ chức: tổng hợp điều phối toàn bộ đường ống, phòng tránh đấu chuyền đường ống; kiểm tra, đối chiếu tên phòng và chức năng sử dụng; kiểm tra, đối chiếu phân khu chống cháy; kiểm tra, đối chiếu xem tài liệu các chuyên ngành nộp và nội dung hiệp thương trong quá trình phối hợp đã được phản ánh trong bản vẽ liên quan chưa; kiểm tra xem bản vẽ thi công của các chuyên ngành có phù hợp nguyên tắc phối hợp thiết kế sơ bộ không; đối chiếu, giải quyết các xung đột có thể có giữa đường ống và cột kết cấu, lanh tô, v.v. trong bản vẽ kiến trúc (hoàn thành trong khoảng thời gian 71%W～76%W của giai đoạn này)",
        "th": "หัวหน้าฝ่ายออกแบบจัด: บูรณาการและประสานท่อทั้งหมดเพื่อป้องกันการชนกันของท่อ ตรวจสอบและทวนสอบชื่อห้องและหน้าที่การใช้งาน ตรวจสอบและทวนสอบการแบ่งเขตอัคคีภัย ตรวจสอบและทวนสอบว่าเอกสารที่ทุกสาขาส่งและเนื้อหาการหารือในกระบวนการประสานงานได้สะท้อนในแบบที่เกี่ยวข้องหรือไม่ ตรวจสอบว่าแบบก่อสร้างของทุกสาขาสอดคล้องกับหลักการประสานงานแบบเบื้องต้นหรือไม่ ตรวจสอบและแก้ไขการชนกันที่อาจเกิดขึ้นระหว่างท่อกับเสาโครงสร้าง ทับหลัง ฯลฯ ในแบบสถาปัตยกรรม (แล้วเสร็จภายในช่วงเวลา 71%W～76%W ของระยะนี้)",
    },
    "水专业提资料（在第三阶段所提资料和管道综合成果基础上）：提建筑——地漏位置和房间坡度、坡向要求，非承重墙预留洞（根据需要提出）；提总图——市政管网（给水、排水、热水等）至本楼接口准确位置，主要室外管道布置及构筑物（化粪池、隔油池、水表井、水泵接合器等）准确位置；提结构——提清全部承重墙、楼板及梁上预留洞准确位置和尺寸，需要结构设计的总图管线固定支架准确位置和推力；提电——完善补充所有本专业控制元器件、控制阀等的位置与控制要求；提智能化——完善补充需远程计量冷热水表具的位置、污水池污水泵的位置（本阶段 76%W～82%W 内完成）": {
        "en": "Materials submitted by the plumbing discipline (on the basis of the materials submitted in the third stage and the results of pipeline integration): to architecture - the positions of floor drains and the requirements of room slopes and slope directions, reserved openings in non-load-bearing walls (proposed as needed); to site planning - the accurate positions of the interfaces of municipal pipeline networks (water supply, drainage, hot water, etc.) to this building, the accurate positions of major outdoor pipeline layouts and structures (septic tanks, grease traps, water meter wells, fire pump connections, etc.); to structural - the accurate positions and dimensions of all reserved openings in load-bearing walls, floor slabs and beams, and the accurate positions and thrusts of fixed supports of site pipelines requiring structural design; to electrical - improve and supplement the positions and control requirements of all control components, control valves, etc. of this discipline; to intelligent systems - improve and supplement the positions of cold and hot water meters requiring remote metering and the positions of sewage pumps in sewage pits (to be completed within 76%W～82%W of this stage)",
        "ja": "水専門が提出資料（第三段階の提出資料と配管総合成果に基づく）：提建築——床ドレン位置と部屋の勾配・勾配方向要求、非耐力壁の予備開口（必要に応じて提出）；提総図——市政管網（給水、排水、給湯など）から本棟への接続正確な位置、主要屋外配管配置及び構築物（化糞池、グリーストラップ、水道メーター井、ポンプ接合器など）の正確な位置；提構造——すべての耐力壁、床版及び梁上の予備開口の正確な位置と寸法、構造設計が必要な総図管線固定支持架の正確な位置と推力；提電——本専門のすべての制御部品、制御弁などの位置と制御要求を改善・補充；提智能化——遠隔計量が必要な冷温水メーターの位置、汚水槽汚水ポンプの位置を改善・補充（本段階 76%W～82%W 内に完了）",
        "ko": "급배수 전문 제출 자료(제3단계 제출 자료와 배관 종합 성과 기반): 건축에 - 바닥 트랩 위치와 실 바닥 기울기·기울기 방향 요구, 비내력벽 예비 구멍(필요 시 제출); 총도에 - 시정 관망(급수, 배수, 온수 등)의 본 건물 접속 정확한 위치, 주요 옥외 배관 배치 및 구조물(정화조, 유분리조, 수도계량기 우물, 펌프 접합기 등)의 정확한 위치; 구조에 - 모든 내력벽, 슬래브 및 보 위 예비 구멍의 정확한 위치와 치수, 구조 설계가 필요한 총도 관선 고정 지지대의 정확한 위치와 추력; 전기에 - 본 전문의 모든 제어 부품, 제어밸브 등의 위치와 제어 요구를 보완·개선; 지능화에 - 원격 계량이 필요한 냉온수 계량기 위치, 오수조 오수 펌프 위치를 보완·개선(이 단계 76%W~82%W 내 완료)",
        "fr": "Documents fournis par la discipline plomberie (sur la base des documents fournis à la troisième phase et des résultats de l'intégration des canalisations) : à l'architecture — positions des siphons de sol et exigences de pentes et de sens des pentes des pièces, réservations dans les murs non porteurs (fournies selon les besoins) ; à l'aménagement du site — positions précises des interfaces des réseaux municipaux (adduction, assainissement, eau chaude, etc.) avec ce bâtiment, positions précises des principaux réseaux extérieurs et ouvrages (fosses septiques, séparateurs à graisses, regards de compteurs d'eau, raccordements de pompes d'incendie, etc.) ; à la structure — positions et dimensions précises de toutes les réservations dans les murs porteurs, dalles et poutres, positions et poussées précises des supports fixes des réseaux du site nécessitant une conception structurelle ; à l'électricité — compléter et améliorer les positions et exigences de commande de tous les composants de commande, vannes de commande, etc. de cette discipline ; aux systèmes intelligents — compléter et améliorer les positions des compteurs d'eau froide et chaude à télémesure et les positions des pompes des cuves à eaux usées (à achever dans la période 76 %W à 82 %W de cette phase)",
        "de": "Vorgelegt von der Sanitärdisziplin (auf der Grundlage der in der dritten Phase vorgelegten Unterlagen und der Ergebnisse der Leitungsintegration): an Architektur — Positionen der Bodenabläufe und Anforderungen an Gefälle und Gefällerichtungen der Räume, Durchbrüche in nicht tragenden Wänden (bei Bedarf); an Standortplanung — die genauen Positionen der Anschlusspunkte der städtischen Netze (Wasserversorgung, Entwässerung, Warmwasser usw.) an dieses Gebäude, die genauen Positionen der Hauptaußenleitungen und Bauwerke (Kläranlagen, Fettabscheider, Wasserzählerschächte, Feuerlöschpumpenanschlüsse usw.); an Tragwerk — die genauen Positionen und Abmessungen aller Durchbrüche in tragenden Wänden, Geschossdecken und Balken sowie die genauen Positionen und Schubkräfte der statisch zu bemessenden Festpunkte der Außenleitungen; an Elektro — die Positionen und Steuerungsanforderungen aller Steuerbauteile, Regelventile usw. dieser Disziplin vervollständigen und verbessern; an Gebäudeautomation — die Positionen der fernzulesenden Kalt- und Warmwasserzähler und die Positionen der Abwasserpumpen in Abwasserschächten vervollständigen und verbessern (innerhalb von 76 %W bis 82 %W dieser Phase abzuschließen)",
        "es": "Documentos aportados por la especialidad de fontanería (sobre la base de los documentos aportados en la tercera fase y los resultados de la integración de canalizaciones): a arquitectura — posiciones de los sumideros de suelo y requisitos de pendientes y sentidos de pendiente de los locales, huecos reservados en muros no portantes (aportados según necesidades); a urbanización — posiciones precisas de los puntos de conexión de las redes municipales (abastecimiento, saneamiento, agua caliente, etc.) con este edificio, posiciones precisas de las principales redes exteriores y obras (fosas sépticas, separadores de grasas, arquetas de contadores de agua, conexiones de bombas de incendio, etc.); a estructuras — posiciones y dimensiones precisas de todos los huecos reservados en muros portantes, forjados y vigas, posiciones y empujes precisos de los soportes fijos de las redes de urbanización que requieren diseño estructural; a electricidad — completar y mejorar las posiciones y requisitos de control de todos los componentes de control, válvulas de control, etc. de esta especialidad; a sistemas inteligentes — completar y mejorar las posiciones de los contadores de agua fría y caliente de telemedida y las posiciones de las bombas de las fosas de aguas residuales (debe completarse dentro del periodo 76%W a 82%W de esta fase)",
        "pt": "Documentos fornecidos pela especialidade de hidráulica (com base nos documentos fornecidos na terceira fase e nos resultados da integração de tubulações): para arquitetura — posições dos ralos de piso e requisitos de declividades e sentidos de declividade dos ambientes, aberturas reservadas em paredes não estruturais (fornecidas conforme necessário); para urbanismo — posições precisas dos pontos de conexão das redes municipais (abastecimento, drenagem, água quente, etc.) com este edifício, posições precisas dos principais traçados externos e obras (fossas sépticas, separadores de gordura, caixas de hidrômetros, conexões de bombas de incêndio, etc.); para estruturas — posições e dimensões precisas de todas as aberturas reservadas em paredes estruturais, lajes e vigas, posições e empuxos precisos dos suportes fixos das redes de urbanismo que exigem projeto estrutural; para elétrica — complementar e aprimorar as posições e requisitos de controle de todos os componentes de controle, válvulas de controle, etc. desta especialidade; para sistemas inteligentes — complementar e aprimorar as posições dos medidores de água fria e quente com telemedição e as posições das bombas das fossas de esgoto (a concluir dentro do período de 76%W a 82%W desta fase)",
        "ru": "Материалы, передаваемые специальностью водоснабжения (на основе материалов, переданных на третьей стадии, и результатов комплексной увязки трубопроводов): архитектуре — расположение трапов и требования к уклонам и направлениям уклонов полов, отверстия в ненесущих стенах (по мере необходимости); генплану — точное расположение точек присоединения городских сетей (водоснабжение, водоотведение, горячее водоснабжение и т.п.) к данному зданию, точное расположение основных наружных сетей и сооружений (септики, жироуловители, колодцы водомеров, пожарные соединительные головки и т.п.); конструкциям — точное расположение и размеры всех отверстий в несущих стенах, плитах перекрытий и балках, точное расположение и распор неподвижных опор наружных сетей, требующих расчёта конструкций; электроснабжению — дополнить и уточнить расположение и требования к управлению всеми контрольно-измерительными приборами, регулирующими клапанами и т.п. данной специальности; интеллектуальным системам — дополнить и уточнить расположение счётчиков холодной и горячей воды с дистанционным учётом и расположение насосов сточных вод в приямках (завершить в период 76%W–82%W данной стадии)",
        "ar": "الوثائق المقدمة من تخصص السباكة (على أساس الوثائق المقدمة في المرحلة الثالثة ونتائج التنسيق الشامل للأنابيب): إلى العمارة — مواقع المصارف الأرضية ومتطلبات الميول واتجاهات الميول للغرف، والفتحات المحجوزة في الجدران غير الحاملة (تُقدم حسب الحاجة)؛ وإلى تخطيط المواقع — المواقع الدقيقة لنقاط ربط شبكات المرافق البلدية (إمداد المياه، الصرف، الماء الساخن وغيرها) بهذا المبنى، والمواقع الدقيقة للشبكات الخارجية الرئيسية والمنشآت (خزانات الصرف الصحي، فواصل الدهون، آبار عدادات المياه، وصلات مضخات الإطفاء وغيرها)؛ وإلى الإنشاءات — المواقع والأبعاد الدقيقة لجميع الفتحات المحجوزة في الجدران الحاملة والألواح والكمرات، والمواقع وقوى الدفع الدقيقة للدعامات الثابتة لخطوط الموقع التي تتطلب تصميمًا إنشائيًا؛ وإلى الكهرباء — استكمال وتحسين مواقع ومتطلبات التحكم لجميع مكونات التحكم وصمامات التحكم وغيرها من هذا التخصص؛ وإلى الأنظمة الذكية — استكمال وتحسين مواقع عدادات الماء البارد والساخن التي تتطلب قياسًا عن بُعد ومواقع مضخات مياه الصرف في الأحواض (يُنجز خلال 76%W～82%W من هذه المرحلة)",
        "it": "Documenti forniti dalla disciplina idrica (sulla base dei documenti forniti nella terza fase e dei risultati dell'integrazione delle tubazioni): ad architettura — posizioni degli scarichi a pavimento e requisiti di pendenze e direzioni di pendenza dei locali, aperture riservate nelle pareti non portanti (fornite secondo necessità); a pianificazione del sito — le posizioni precise dei punti di collegamento delle reti comunali (acquedotto, fognatura, acqua calda, ecc.) con questo edificio, le posizioni precise dei principali tracciati esterni e delle opere (fosse settiche, separatori di grassi, pozzetti contatori acqua, attacchi pompe antincendio, ecc.); a strutture — le posizioni e dimensioni precise di tutte le aperture riservate nelle pareti portanti, nei solai e nelle travi, le posizioni e spinte precise dei supporti fissi delle reti di sito che richiedono progettazione strutturale; a elettrici — integrare e migliorare le posizioni e i requisiti di controllo di tutti i componenti di controllo, valvole di controllo, ecc. di questa disciplina; a sistemi intelligenti — integrare e migliorare le posizioni dei contatori di acqua fredda e calda con telelettura e le posizioni delle pompe di sollevamento delle vasche di acque reflue (da completare entro il periodo 76%W～82%W di questa fase)",
        "vi": "Tài liệu chuyên ngành cấp thoát nước nộp (trên cơ sở tài liệu nộp ở giai đoạn 3 và thành quả tổng hợp đường ống): gửi kiến trúc — vị trí lỗ thu nước sàn và yêu cầu độ dốc, hướng dốc của phòng, lỗ chờ tường không chịu lực (nêu theo nhu cầu); gửi tổng mặt bằng — vị trí chính xác điểm đấu nối mạng hạ tầng (cấp nước, thoát nước, nước nóng, v.v.) vào công trình, vị trí chính xác bố trí đường ống ngoài nhà chính và công trình (bể tự hoại, bể tách dầu mỡ, giếng đồng hồ nước, đầu nối máy bơm chữa cháy, v.v.); gửi kết cấu — vị trí và kích thước chính xác toàn bộ lỗ chờ trên tường chịu lực, sàn và dầm, vị trí và lực đẩy chính xác gối đỡ cố định đường ống tổng mặt bằng cần kết cấu thiết kế; gửi điện — hoàn thiện, bổ sung vị trí và yêu cầu điều khiển của toàn bộ linh kiện điều khiển, van điều khiển, v.v. của chuyên ngành mình; gửi hệ thống thông minh — hoàn thiện, bổ sung vị trí đồng hồ nước nóng lạnh cần đo xa và vị trí bơm nước thải trong bể nước thải (hoàn thành trong 76%W～82%W của giai đoạn này)",
        "th": "เอกสารที่สาขาประปาส่ง (บนพื้นฐานเอกสารที่ส่งในระยะที่สามและผลการบูรณาการท่อ): ส่งให้สถาปัตยกรรม — ตำแหน่งท่อระบายน้ำพื้นและข้อกำหนดความลาดเอียงและทิศทางลาดเอียงของห้อง ช่องเผื่อในผนังไม่รับน้ำหนัก (ส่งตามความจำเป็น) ส่งให้การวางผังบริเวณ — ตำแหน่งที่แม่นยำของจุดเชื่อมต่อเครือข่ายสาธารณูปโภค (ประปา ท่อระบายน้ำ น้ำร้อน ฯลฯ) เข้าสู่อาคาร ตำแหน่งที่แม่นยำของการวางท่อภายนอกหลักและโครงสร้าง (ถังบำบัด ถังดักไขมัน บ่อมาตรวัดน้ำ ข้อต่อปั๊มดับเพลิง ฯลฯ) ส่งให้โครงสร้าง — ตำแหน่งและขนาดที่แม่นยำของช่องเผื่อทั้งหมดในผนังรับน้ำหนัก พื้น และคาน ตำแหน่งและแรงดันที่แม่นยำของขายึดท่อสายวางผังที่ต้องออกแบบโครงสร้าง ส่งให้ไฟฟ้า — ปรับปรุงและเพิ่มเติมตำแหน่งและข้อกำหนดการควบคุมของส่วนควบคุม วาล์วควบคุม ฯลฯ ทั้งหมดของสาขานี้ ส่งให้ระบบอัจฉริยะ — ปรับปรุงและเพิ่มเติมตำแหน่งมาตรวัดน้ำเย็นน้ำร้อนที่ต้องวัดระยะไกลและตำแหน่งปั๊มน้ำเสียในบ่อน้ำเสีย (แล้วเสร็จภายใน 76%W～82%W ของระยะนี้)",
    },
    "暖专业提资料（在第三阶段所提资料和管道综合成果基础上）：提建筑——确认和补齐全部室外新风口和排风口的尺寸和位置，非承重墙预留洞（根据需要提出）；提总图——市政管网（热力、燃气）至本楼接口准确位置，主要室外管道布置及构筑物（阀门井等）准确位置；提结构——提清全部承重墙、楼板及梁上预留洞准确位置和尺寸，需要结构设计的总图管线固定支架准确位置和推力；提电——完善补充所有本专业控制元器件、控制阀等的位置与控制要求；提智能化——提出空调通风设备自动化控制要求（本阶段 76%W～82%W 内完成）": {
        "en": "Materials submitted by the HVAC discipline (on the basis of the materials submitted in the third stage and the results of pipeline integration): to architecture - confirm and complete the dimensions and positions of all outdoor fresh air outlets and exhaust outlets, reserved openings in non-load-bearing walls (proposed as needed); to site planning - the accurate positions of the interfaces of municipal pipeline networks (heating, gas) to this building, the accurate positions of major outdoor pipeline layouts and structures (valve wells, etc.); to structural - the accurate positions and dimensions of all reserved openings in load-bearing walls, floor slabs and beams, and the accurate positions and thrusts of fixed supports of site pipelines requiring structural design; to electrical - improve and supplement the positions and control requirements of all control components, control valves, etc. of this discipline; to intelligent systems - propose the requirements for automatic control of air-conditioning and ventilation equipment (to be completed within 76%W～82%W of this stage)",
        "ja": "暖専門が提出資料（第三段階の提出資料と配管総合成果に基づく）：提建築——すべての室外新風口と排風口の寸法と位置を確認・補完、非耐力壁の予備開口（必要に応じて提出）；提総図——市政管網（熱力、ガス）から本棟への接続正確な位置、主要屋外配管配置及び構築物（バルブ井など）の正確な位置；提構造——すべての耐力壁、床版及び梁上の予備開口の正確な位置と寸法、構造設計が必要な総図管線固定支持架の正確な位置と推力；提電——本専門のすべての制御部品、制御弁などの位置と制御要求を改善・補充；提智能化——空調換気設備の自動制御要求を提出（本段階 76%W～82%W 内に完了）",
        "ko": "난방 전문 제출 자료(제3단계 제출 자료와 배관 종합 성과 기반): 건축에 - 모든 옥외 신풍구와 배풍구의 치수와 위치를 확인·보완, 비내력벽 예비 구멍(필요 시 제출); 총도에 - 시정 관망(열력, 가스)의 본 건물 접속 정확한 위치, 주요 옥외 배관 배치 및 구조물(밸브 우물 등)의 정확한 위치; 구조에 - 모든 내력벽, 슬래브 및 보 위 예비 구멍의 정확한 위치와 치수, 구조 설계가 필요한 총도 관선 고정 지지대의 정확한 위치와 추력; 전기에 - 본 전문의 모든 제어 부품, 제어밸브 등의 위치와 제어 요구를 보완·개선; 지능화에 - 공조 환기 설비의 자동 제어 요구를 제출(이 단계 76%W~82%W 내 완료)",
        "fr": "Documents fournis par la discipline CVC (sur la base des documents fournis à la troisième phase et des résultats de l'intégration des canalisations) : à l'architecture — confirmer et compléter les dimensions et positions de toutes les bouches d'air neuf et d'extraction extérieures, réservations dans les murs non porteurs (fournies selon les besoins) ; à l'aménagement du site — positions précises des interfaces des réseaux municipaux (chauffage, gaz) avec ce bâtiment, positions précises des principaux réseaux extérieurs et ouvrages (regards de vannes, etc.) ; à la structure — positions et dimensions précises de toutes les réservations dans les murs porteurs, dalles et poutres, positions et poussées précises des supports fixes des réseaux du site nécessitant une conception structurelle ; à l'électricité — compléter et améliorer les positions et exigences de commande de tous les composants de commande, vannes de commande, etc. de cette discipline ; aux systèmes intelligents — proposer les exigences de commande automatique des équipements de climatisation et de ventilation (à achever dans la période 76 %W à 82 %W de cette phase)",
        "de": "Vorgelegt von der Disziplin Heizung-Lüftung (auf der Grundlage der in der dritten Phase vorgelegten Unterlagen und der Ergebnisse der Leitungsintegration): an Architektur — die Abmessungen und Positionen aller außenliegenden Frischluft- und Abluftöffnungen bestätigen und vervollständigen, Durchbrüche in nicht tragenden Wänden (bei Bedarf); an Standortplanung — die genauen Positionen der Anschlusspunkte der städtischen Netze (Fernwärme, Gas) an dieses Gebäude, die genauen Positionen der Hauptaußenleitungen und Bauwerke (Absperrschieberschächte usw.); an Tragwerk — die genauen Positionen und Abmessungen aller Durchbrüche in tragenden Wänden, Geschossdecken und Balken sowie die genauen Positionen und Schubkräfte der statisch zu bemessenden Festpunkte der Außenleitungen; an Elektro — die Positionen und Steuerungsanforderungen aller Steuerbauteile, Regelventile usw. dieser Disziplin vervollständigen und verbessern; an Gebäudeautomation — die Anforderungen an die automatische Steuerung der Klima- und Lüftungsanlagen vorlegen (innerhalb von 76 %W bis 82 %W dieser Phase abzuschließen)",
        "es": "Documentos aportados por la especialidad de climatización (sobre la base de los documentos aportados en la tercera fase y los resultados de la integración de canalizaciones): a arquitectura — confirmar y completar las dimensiones y posiciones de todas las tomas de aire nuevo y salidas de extracción exteriores, huecos reservados en muros no portantes (aportados según necesidades); a urbanización — posiciones precisas de los puntos de conexión de las redes municipales (calefacción, gas) con este edificio, posiciones precisas de las principales redes exteriores y obras (arquetas de válvulas, etc.); a estructuras — posiciones y dimensiones precisas de todos los huecos reservados en muros portantes, forjados y vigas, posiciones y empujes precisos de los soportes fijos de las redes de urbanización que requieren diseño estructural; a electricidad — completar y mejorar las posiciones y requisitos de control de todos los componentes de control, válvulas de control, etc. de esta especialidad; a sistemas inteligentes — proponer los requisitos de control automático de los equipos de climatización y ventilación (debe completarse dentro del periodo 76%W a 82%W de esta fase)",
        "pt": "Documentos fornecidos pela especialidade de climatização (com base nos documentos fornecidos na terceira fase e nos resultados da integração de tubulações): para arquitetura — confirmar e completar as dimensões e posições de todas as entradas de ar novo e saídas de exaustão externas, aberturas reservadas em paredes não estruturais (fornecidas conforme necessário); para urbanismo — posições precisas dos pontos de conexão das redes municipais (aquecimento, gás) com este edifício, posições precisas dos principais traçados externos e obras (caixas de válvulas, etc.); para estruturas — posições e dimensões precisas de todas as aberturas reservadas em paredes estruturais, lajes e vigas, posições e empuxos precisos dos suportes fixos das redes de urbanismo que exigem projeto estrutural; para elétrica — complementar e aprimorar as posições e requisitos de controle de todos os componentes de controle, válvulas de controle, etc. desta especialidade; para sistemas inteligentes — propor os requisitos de controle automático dos equipamentos de climatização e ventilação (a concluir dentro do período de 76%W a 82%W desta fase)",
        "ru": "Материалы, передаваемые специальностью отопления и вентиляции (на основе материалов, переданных на третьей стадии, и результатов комплексной увязки трубопроводов): архитектуре — подтвердить и дополнить размеры и расположение всех наружных воздухозаборов и вытяжных отверстий, отверстия в ненесущих стенах (по мере необходимости); генплану — точное расположение точек присоединения городских сетей (теплоснабжение, газ) к данному зданию, точное расположение основных наружных сетей и сооружений (задвижки в колодцах и т.п.); конструкциям — точное расположение и размеры всех отверстий в несущих стенах, плитах перекрытий и балках, точное расположение и распор неподвижных опор наружных сетей, требующих расчёта конструкций; электроснабжению — дополнить и уточнить расположение и требования к управлению всеми контрольно-измерительными приборами, регулирующими клапанами и т.п. данной специальности; интеллектуальным системам — представить требования к автоматическому управлению оборудованием кондиционирования и вентиляции (завершить в период 76%W–82%W данной стадии)",
        "ar": "الوثائق المقدمة من تخصص التدفئة والتهوية (على أساس الوثائق المقدمة في المرحلة الثالثة ونتائج التنسيق الشامل للأنابيب): إلى العمارة — تأكيد واستكمال أبعاد ومواقع جميع مداخل الهواء الجديد ومخارج الهواء العادم الخارجية، والفتحات المحجوزة في الجدران غير الحاملة (تُقدم حسب الحاجة)؛ وإلى تخطيط المواقع — المواقع الدقيقة لنقاط ربط شبكات المرافق البلدية (الحرارة، الغاز) بهذا المبنى، والمواقع الدقيقة للشبكات الخارجية الرئيسية والمنشآت (آبار الصمامات وغيرها)؛ وإلى الإنشاءات — المواقع والأبعاد الدقيقة لجميع الفتحات المحجوزة في الجدران الحاملة والألواح والكمرات، والمواقع وقوى الدفع الدقيقة للدعامات الثابتة لخطوط الموقع التي تتطلب تصميمًا إنشائيًا؛ وإلى الكهرباء — استكمال وتحسين مواقع ومتطلبات التحكم لجميع مكونات التحكم وصمامات التحكم وغيرها من هذا التخصص؛ وإلى الأنظمة الذكية — تقديم متطلبات التحكم الآلي لمعدات التكييف والتهوية (يُنجز خلال 76%W～82%W من هذه المرحلة)",
        "it": "Documenti forniti dalla disciplina HVAC (sulla base dei documenti forniti nella terza fase e dei risultati dell'integrazione delle tubazioni): ad architettura — confermare e completare le dimensioni e posizioni di tutte le prese d'aria nuova e le bocchette di estrazione esterne, aperture riservate nelle pareti non portanti (fornite secondo necessità); a pianificazione del sito — le posizioni precise dei punti di collegamento delle reti comunali (teleriscaldamento, gas) con questo edificio, le posizioni precise dei principali tracciati esterni e delle opere (pozzetti valvole, ecc.); a strutture — le posizioni e dimensioni precise di tutte le aperture riservate nelle pareti portanti, nei solai e nelle travi, le posizioni e spinte precise dei supporti fissi delle reti di sito che richiedono progettazione strutturale; a elettrici — integrare e migliorare le posizioni e i requisiti di controllo di tutti i componenti di controllo, valvole di controllo, ecc. di questa disciplina; a sistemi intelligenti — proporre i requisiti di controllo automatico delle apparecchiature di climatizzazione e ventilazione (da completare entro il periodo 76%W～82%W di questa fase)",
        "vi": "Tài liệu chuyên ngành điều hòa nộp (trên cơ sở tài liệu nộp ở giai đoạn 3 và thành quả tổng hợp đường ống): gửi kiến trúc — xác nhận và bổ sung đầy đủ kích thước, vị trí các miệng gió tươi và gió thải ngoài trời, lỗ chờ tường không chịu lực (nêu theo nhu cầu); gửi tổng mặt bằng — vị trí chính xác điểm đấu nối mạng hạ tầng (nhiệt, gas) vào công trình, vị trí chính xác bố trí đường ống ngoài nhà chính và công trình (giếng van, v.v.); gửi kết cấu — vị trí và kích thước chính xác toàn bộ lỗ chờ trên tường chịu lực, sàn và dầm, vị trí và lực đẩy chính xác gối đỡ cố định đường ống tổng mặt bằng cần kết cấu thiết kế; gửi điện — hoàn thiện, bổ sung vị trí và yêu cầu điều khiển của toàn bộ linh kiện điều khiển, van điều khiển, v.v. của chuyên ngành mình; gửi hệ thống thông minh — nêu yêu cầu điều khiển tự động thiết bị điều hòa thông gió (hoàn thành trong 76%W～82%W của giai đoạn này)",
        "th": "เอกสารที่สาขาระบบปรับอากาศส่ง (บนพื้นฐานเอกสารที่ส่งในระยะที่สามและผลการบูรณาการท่อ): ส่งให้สถาปัตยกรรม — ยืนยันและเติมเต็มขนาดและตำแหน่งของช่องลมเข้าและลมออกภายนอกทั้งหมด ช่องเผื่อในผนังไม่รับน้ำหนัก (ส่งตามความจำเป็น) ส่งให้การวางผังบริเวณ — ตำแหน่งที่แม่นยำของจุดเชื่อมต่อเครือข่ายสาธารณูปโภค (ความร้อน แก๊ส) เข้าสู่อาคาร ตำแหน่งที่แม่นยำของการวางท่อภายนอกหลักและโครงสร้าง (บ่อวาล์ว ฯลฯ) ส่งให้โครงสร้าง — ตำแหน่งและขนาดที่แม่นยำของช่องเผื่อทั้งหมดในผนังรับน้ำหนัก พื้น และคาน ตำแหน่งและแรงดันที่แม่นยำของขายึดท่อสายวางผังที่ต้องออกแบบโครงสร้าง ส่งให้ไฟฟ้า — ปรับปรุงและเพิ่มเติมตำแหน่งและข้อกำหนดการควบคุมของส่วนควบคุม วาล์วควบคุม ฯลฯ ทั้งหมดของสาขานี้ ส่งให้ระบบอัจฉริยะ — เสนอข้อกำหนดการควบคุมอัตโนมัติของอุปกรณ์ปรับอากาศและระบายอากาศ (แล้วเสร็จภายใน 76%W～82%W ของระยะนี้)",
    },
    "电专业提资料（在第三阶段所提资料和管道综合成果基础上）：提建筑——各房间门开启方向的要求，各消防控制按钮的位置要求；提总图——市政管线（电力、通讯等）至本楼接口准确位置，主要室外管线布置及构筑物（检修井等）准确位置；提结构——提清全部承重墙、楼板及梁上预留洞准确位置和尺寸；提智能化——完善补充需远程计量电表的位置，完善和补充需由建筑设备管理系统监视和控制的配电箱位置（本阶段 76%W～82%W 内完成）": {
        "en": "The electrical discipline submits materials (on the basis of the materials submitted in the third stage and the results of pipe comprehensive coordination): to architecture — requirements for the opening direction of doors of each room and the position of fire-fighting control buttons; to general plan — accurate positions of the interfaces of municipal pipelines (power, communication, etc.) to this building, and accurate positions of main outdoor pipeline layouts and structures (inspection wells, etc.); to structure — clearly provide the accurate positions and dimensions of all reserved openings on load-bearing walls, floor slabs and beams; to intelligent building — supplement the positions of meters requiring remote metering, and supplement the positions of distribution boxes requiring monitoring and control by the building equipment management system (to be completed within 76%W~82%W of this stage)",
        "ja": "電気専門が資料を提出する（第三段階で提出した資料と配管総合成果を基礎とする）：建築に対し——各室扉の開き方向の要求、各消防制御ボタンの位置要求；総図に対し——市政管線（電力、通信等）から本建物への接続口の正確な位置、主要な屋外管線配置と構造物（点検井等）の正確な位置；構造に対し——全ての耐力壁、スラブ及び梁上の予留孔の正確な位置と寸法を明示；智能化に対し——遠隔計量が必要な電表の位置を整備・補充し、建築設備管理システムにより監視・制御される配電箱の位置を整備・補充する（本段階の76%W～82%W以内に完了）",
        "ko": "전기 전문이 자료를 제출한다(제3단계에서 제출한 자료와 배관 종합 성과를 기초로 한다)：건축에 대해——각 실 문 개폐 방향 요구, 각 소방 제어 버튼 위치 요구；총도에 대해——시정 관로(전력, 통신 등)의 본 건물 접속구 정확한 위치, 주요 옥외 관로 배치 및 구조물(점검정 등) 정확한 위치；구조에 대해——모든 내력벽, 슬래브 및 보의 예비 구멍 정확한 위치와 치수를 명확히 제시；지능화에 대해——원격 계량이 필요한 전기계량기 위치를 정비·보충하고, 건축 설비 관리 시스템의 감시·제어가 필요한 배전함 위치를 정비·보충한다(본 단계 76%W～82%W 내 완료)",
        "fr": "La spécialité électrique soumet les documents (sur la base des documents soumis à la troisième étape et des résultats de la coordination des réseaux) : à l'architecture — exigences relatives au sens d'ouverture des portes de chaque pièce et à la position des boutons de commande de sécurité incendie ; au plan général — positions précises des interfaces des réseaux municipaux (électricité, communications, etc.) vers ce bâtiment, et positions précises des principaux réseaux extérieurs et ouvrages (regards, etc.) ; à la structure — fournir clairement les positions et dimensions précises de toutes les réservations sur les murs porteurs, les dalles et les poutres ; à l'intelligence du bâtiment — compléter les positions des compteurs nécessitant un relevé à distance, et compléter les positions des armoires de distribution devant être surveillées et contrôlées par le système de gestion des équipements du bâtiment (à terminer dans les 76%W~82%W de cette étape)",
        "de": "Die Fachrichtung Elektrotechnik reicht Unterlagen ein (auf der Grundlage der in der dritten Phase eingereichten Unterlagen und der Ergebnisse der Leitungsgesamtkoordination): an die Architektur — Anforderungen an die Öffnungsrichtung der Türen jedes Raums und an die Position der Brandschutz-Bedientasten; an den Lageplan — genaue Positionen der Anschlusspunkte der kommunalen Leitungen (Strom, Kommunikation usw.) zu diesem Gebäude sowie genaue Positionen der wichtigsten Außenleitungen und Bauwerke (Kontrollschächte usw.); an die Tragwerksplanung — eindeutige Angabe der genauen Positionen und Abmessungen aller Vorhalteöffnungen in tragenden Wänden, Deckenplatten und Trägern; an die Gebäudeintelligenz — Ergänzung der Positionen der Zähler, die eine Fernmessung erfordern, und Ergänzung der Positionen der Verteilerschränke, die vom Gebäudemanagementsystem überwacht und gesteuert werden müssen (innerhalb von 76%W~82%W dieser Phase abzuschließen)",
        "es": "La especialidad eléctrica presenta materiales (sobre la base de los materiales presentados en la tercera etapa y los resultados de la coordinación integral de tuberías): a arquitectura — requisitos sobre la dirección de apertura de las puertas de cada habitación y la posición de los botones de control contra incendios; al plano general — posiciones precisas de las interfaces de las redes municipales (electricidad, comunicaciones, etc.) hacia este edificio, y posiciones precisas de los principales trazados exteriores de redes y estructuras (pozos de inspección, etc.); a estructura — proporcionar claramente las posiciones y dimensiones precisas de todas las reservas en muros portantes, losas y vigas; a inteligencia del edificio — completar y complementar las posiciones de los contadores que requieren medición remota, y completar y complementar las posiciones de los cuadros de distribución que deben ser supervisados y controlados por el sistema de gestión de equipos del edificio (debe completarse dentro del 76%W~82%W de esta etapa)",
        "pt": "A especialidade elétrica apresenta materiais (com base nos materiais apresentados na terceira etapa e nos resultados da coordenação integral de tubulações): à arquitetura — requisitos sobre a direção de abertura das portas de cada ambiente e a posição dos botões de controle de incêndio; ao plano geral — posições precisas das interfaces das redes municipais (eletricidade, comunicações, etc.) para este edifício e posições precisas dos principais traçados externos de redes e estruturas (poços de inspeção, etc.); à estrutura — fornecer claramente as posições e dimensões precisas de todas as reservas em paredes estruturais, lajes e vigas; à inteligência predial — complementar as posições dos medidores que exigem medição remota e complementar as posições dos quadros de distribuição que devem ser supervisionados e controlados pelo sistema de gestão de equipamentos do edifício (a concluir dentro de 76%W~82%W desta etapa)",
        "ru": "Электротехническая специальность представляет материалы (на основе материалов, представленных на третьем этапе, и результатов комплексной координации трубопроводов): архитектуре — требования к направлению открывания дверей каждого помещения и расположению кнопок пожарного управления; генеральному плану — точные места присоединения городских сетей (электроснабжение, связь и т.п.) к данному зданию, точное расположение основных наружных сетей и сооружений (смотровых колодцев и т.п.); конструкции — четко указать точные места и размеры всех отверстий в несущих стенах, плитах перекрытия и балках; интеллектуальному зданию — дополнить места установки счетчиков, требующих дистанционного учета, и дополнить места установки распределительных щитов, подлежащих контролю и управлению системой управления инженерным оборудованием здания (завершить в пределах 76%W~82%W данного этапа)",
        "ar": "يقدم تخصص الكهرباء المواد (بناءً على المواد المقدمة في المرحلة الثالثة ونتائج التنسيق الشامل للخطوط): إلى العمارة — متطلبات اتجاه فتح أبواب كل غرفة ومواقع أزرار التحكم في الإطفاء؛ إلى المخطط العام — المواقع الدقيقة لواجهات الشبكات البلدية (الكهرباء والاتصالات وغيرها) نحو هذا المبنى والمواقع الدقيقة لخطوط الشبكات الخارجية الرئيسية والمنشآت (آبار التفتيش وغيرها)؛ إلى الإنشاءات — توضيح المواقع والأبعاد الدقيقة لجميع الفتحات المحجوزة في الجدران الحاملة والأسقف والكمرات؛ إلى المباني الذكية — استكمال مواقع العدادات التي تتطلب قياسًا عن بُعد واستكمال مواقع لوحات التوزيع التي يجب مراقبتها والتحكم بها بواسطة نظام إدارة معدات المبنى (يُكمل خلال 76%W~82%W من هذه المرحلة)",
        "it": "La disciplina elettrica presenta i materiali (sulla base dei materiali presentati nella terza fase e dei risultati del coordinamento integrale delle tubazioni): all'architettura — requisiti sulla direzione di apertura delle porte di ogni locale e sulla posizione dei pulsanti di comando antincendio; al piano generale — posizioni precise delle interfacce delle reti municipali (elettricità, comunicazioni, ecc.) verso questo edificio e posizioni precise dei principali tracciati esterni delle reti e delle opere (pozzetti di ispezione, ecc.); alla struttura — indicare chiaramente le posizioni e dimensioni precise di tutte le riserve su muri portanti, solai e travi; all'edificio intelligente — integrare e completare le posizioni dei contatori che richiedono misurazione remota e completare le posizioni dei quadri di distribuzione che devono essere monitorati e controllati dal sistema di gestione delle apparecchiature dell'edificio (da completare entro il 76%W~82%W di questa fase)",
        "vi": "Chuyên ngành điện trình tài liệu (trên cơ sở tài liệu đã trình ở giai đoạn thứ ba và kết quả tổng hợp đường ống): cho kiến trúc — yêu cầu về hướng mở cửa của từng phòng, vị trí các nút điều khiển phòng cháy; cho tổng mặt bằng — vị trí chính xác các điểm đấu nối của đường ống đô thị (điện, thông tin...) vào tòa nhà này, vị trí chính xác các tuyến ống ngoài trời chính và công trình (hố kiểm tra...); cho kết cấu — nêu rõ vị trí và kích thước chính xác của tất cả lỗ chừa trên tường chịu lực, sàn và dầm; cho trí tuệ hóa — bổ sung vị trí các công tơ cần đo xa, bổ sung vị trí các tủ phân phối cần được giám sát và điều khiển bởi hệ thống quản lý thiết bị tòa nhà (hoàn thành trong 76%W～82%W của giai đoạn này)",
        "th": "สาขาไฟฟ้ายื่นเอกสาร (บนพื้นฐานของเอกสารที่ยื่นในขั้นตอนที่สามและผลการรวมท่อ): ให้สถาปัตยกรรม — ข้อกำหนดทิศทางการเปิดประตูแต่ละห้องและตำแหน่งปุ่มควบคุมอัคคีภัย ให้ผังรวม — ตำแหน่งที่แน่นอนของจุดเชื่อมต่อระบบสาธารณูปโภค (ไฟฟ้า โทรคมนาคม ฯลฯ) เข้าสู่อาคารนี้ ตำแหน่งที่แน่นอนของแนวท่อภายนอกหลักและสิ่งก่อสร้าง (บ่อตรวจสอบ ฯลฯ) ให้โครงสร้าง — ระบุตำแหน่งและขนาดที่แน่นอนของช่องเว้นทั้งหมดบนผนังรับน้ำหนัก พื้นและคาน ให้อัจฉริยภาพอาคาร — เพิ่มเติมตำแหน่งมิเตอร์ที่ต้องวัดระยะไกล และเพิ่มเติมตำแหน่งตู้จ่ายไฟที่ต้องเฝ้าติดตามและควบคุมโดยระบบจัดการอุปกรณ์อาคาร (แล้วเสร็จภายใน 76%W～82%W ของขั้นตอนนี้)",
    },
    "智能化专业提资料：对电气专业提出平面上数据插座、有线电视插座等需预留~220V电源的位置（本阶段 76%W～82%W 内完成）": {
        "en": "The intelligent building discipline submits materials: to the electrical discipline, propose the positions where ~220V power supply shall be reserved for data sockets, cable TV sockets, etc. on the plan (to be completed within 76%W~82%W of this stage)",
        "ja": "智能化専門が資料を提出する：電気専門に対し、平面図上でデータコンセント、ケーブルテレビコンセント等の~220V電源を予留すべき位置を提示する（本段階の76%W～82%W以内に完了）",
        "ko": "지능화 전문이 자료를 제출한다：전기 전문에 대해 평면상 데이터 소켓, 유선TV 소켓 등 ~220V 전원을 예비(預留)해야 할 위치를 제시한다(본 단계 76%W～82%W 내 완료)",
        "fr": "La spécialité d'intelligence du bâtiment soumet les documents : proposer à la spécialité électrique les emplacements où une alimentation ~220V doit être réservée pour les prises de données, les prises de télévision par câble, etc. sur le plan (à terminer dans les 76%W~82%W de cette étape)",
        "de": "Die Fachrichtung Gebäudeintelligenz reicht Unterlagen ein: der Fachrichtung Elektrotechnik werden auf dem Plan die Positionen mitgeteilt, an denen ~220V-Stromversorgung für Datenanschlussdosen, Kabelfernseh-Anschlussdosen usw. vorzusehen ist (innerhalb von 76%W~82%W dieser Phase abzuschließen)",
        "es": "La especialidad de inteligencia del edificio presenta materiales: proponer a la especialidad eléctrica las posiciones donde se debe reservar la alimentación de ~220V para tomas de datos, tomas de televisión por cable, etc. en el plano (debe completarse dentro del 76%W~82%W de esta etapa)",
        "pt": "A especialidade de inteligência predial apresenta materiais: propor à especialidade elétrica as posições onde se deve reservar a alimentação ~220V para tomadas de dados, tomadas de TV a cabo, etc. na planta (a concluir dentro de 76%W~82%W desta etapa)",
        "ru": "Специальность интеллектуального здания представляет материалы: указать электротехнической специальности на плане места, где необходимо предусмотреть резерв электропитания ~220 В для розеток для данных, розеток кабельного телевидения и т.п. (завершить в пределах 76%W~82%W данного этапа)",
        "ar": "يقدم تخصص المباني الذكية المواد: اقتراح على تخصص الكهرباء المواضع التي يجب فيها حجز مصدر طاقة ~220 فولت لمقابس البيانات ومقابس التلفزيون الكبلي وغيرها على المخطط (يُكمل خلال 76%W~82%W من هذه المرحلة)",
        "it": "La disciplina di edificio intelligente presenta i materiali: proporre alla disciplina elettrica le posizioni in cui riservare l'alimentazione ~220V per prese dati, prese TV via cavo, ecc. in pianta (da completare entro il 76%W~82%W di questa fase)",
        "vi": "Chuyên ngành trí tuệ hóa tòa nhà trình tài liệu: đề xuất cho chuyên ngành điện các vị trí cần dự trữ nguồn điện ~220V cho ổ cắm dữ liệu, ổ cắm truyền hình cáp... trên mặt bằng (hoàn thành trong 76%W～82%W của giai đoạn này)",
        "th": "สาขาอัจฉริยภาพอาคารยื่นเอกสาร: เสนอต่อสาขาไฟฟ้าเกี่ยวกับตำแหน่งที่ต้องสำรองแหล่งจ่ายไฟ ~220V สำหรับเต้ารับข้อมูล เต้ารับเคเบิลทีวี ฯลฯ บนแปลน (แล้วเสร็จภายใน 76%W～82%W ของขั้นตอนนี้)",
    },
    "建筑专业提出本版建筑图纸及相关的资料，应反映上述各过程配合后的综合成果，作为各专业最终出图用的建筑底图；本版图纸可以采用电子文件，但与第三阶段（第二版建筑作业图）有修改之处应在电子文件中做出明确标记、标识或书面的文字说明；电子文件按图中信息分层管理；各专业接收后应根据本专业图纸表达特点进行必要修整，删除与本专业图纸表达无关的图形、文字、尺寸等": {
        "en": "The architecture discipline issues the current version of architectural drawings and related materials, which shall reflect the comprehensive results after the coordination of the above processes, and serve as the architectural base drawings for final drawing production of each discipline; the current version may be in electronic files, but modifications from the third stage (second-version architectural working drawings) shall be clearly marked, identified or stated in writing in the electronic files; the electronic files shall be managed by layers according to the information in the drawings; after receiving, each discipline shall make necessary adjustments according to the expression characteristics of its own drawings and delete the graphics, texts, dimensions, etc. irrelevant to its own drawing expression",
        "ja": "建築専門は本版の建築図面及び関連資料を提出し、上記各プロセスの協働後の総合成果を反映し、各専門の最終出図用の建築ベース図とする；本版図面は電子ファイルでもよいが、第三段階（第二版建築作業図）からの変更箇所は電子ファイル内に明確なマーク、表示又は書面による説明を付すこと；電子ファイルは図内情報に応じてレイヤー管理する；各専門は受領後、自専門の図面表現の特徴に応じて必要な修正を行い、自専門の図面表現に関係のない図形、文字、寸法などを削除する",
        "ko": "건축 전문은 본 버전의 건축 도면 및 관련 자료를 제출하며, 상기 각 과정의 협동 후 종합 성과를 반영하여 각 전문의 최종 출도용 건축 베이스 도면으로 한다；본 버전 도면은 전자 파일로 할 수 있으나, 제3단계(제2판 건축 작업도)와 달라진 부분은 전자 파일에 명확한 표시, 식별 또는 서면 설명을 명기해야 한다；전자 파일은 도면 내 정보에 따라 레이어로 관리한다；각 전문은 수령 후 자체 전문 도면 표현 특성에 따라 필요한 수정을 하고, 자체 전문 도면 표현과 무관한 도형, 문자, 치수 등을 삭제한다",
        "fr": "La spécialité d'architecture présente la version actuelle des plans d'architecture et des documents associés, qui doit refléter les résultats complets après la coordination des processus ci-dessus et servir de fond de plan d'architecture pour la production finale des plans de chaque spécialité ; cette version peut être en fichiers électroniques, mais les modifications par rapport à la troisième étape (deuxième version des plans de travail d'architecture) doivent être clairement marquées, identifiées ou expliquées par écrit dans les fichiers électroniques ; les fichiers électroniques sont gérés par couches selon les informations des plans ; après réception, chaque spécialité doit procéder aux ajustements nécessaires selon les caractéristiques d'expression de ses propres plans et supprimer les graphiques, textes, cotes, etc. sans rapport avec sa propre expression",
        "de": "Die Fachrichtung Architektur legt die aktuelle Fassung der Architekturpläne und zugehörigen Unterlagen vor, die die umfassenden Ergebnisse nach der Koordination der oben genannten Prozesse widerspiegeln und als Architektur-Grundlage für die endgültige Planerstellung aller Fachrichtungen dienen sollen; diese Fassung kann in elektronischen Dateien erfolgen, Änderungen gegenüber der dritten Phase (zweite Fassung der Architektur-Arbeitspläne) müssen jedoch in den elektronischen Dateien deutlich markiert, gekennzeichnet oder schriftlich erläutert werden; die elektronischen Dateien sind entsprechend der Informationen im Plan in Ebenen zu verwalten; nach dem Empfang hat jede Fachrichtung die erforderlichen Anpassungen gemäß den Darstellungseigenheiten ihrer eigenen Pläne vorzunehmen und grafische Darstellungen, Texte, Maße usw., die für die eigene Darstellung ohne Belang sind, zu entfernen",
        "es": "La especialidad de arquitectura presenta la versión actual de los planos de arquitectura y materiales relacionados, que debe reflejar los resultados integrales después de la coordinación de los procesos mencionados, y servir como planos base de arquitectura para la producción final de planos de cada especialidad; esta versión puede ser en archivos electrónicos, pero las modificaciones respecto a la tercera etapa (segunda versión de los planos de trabajo de arquitectura) deben estar claramente marcadas, identificadas o explicadas por escrito en los archivos electrónicos; los archivos electrónicos se gestionan por capas según la información de los planos; después de la recepción, cada especialidad debe realizar los ajustes necesarios según las características de expresión de sus propios planos y eliminar los gráficos, textos, cotas, etc. que no estén relacionados con su propia expresión",
        "pt": "A especialidade de arquitetura apresenta a versão atual dos desenhos de arquitetura e materiais relacionados, que deve refletir os resultados abrangentes após a coordenação dos processos acima e servir como desenhos-base de arquitetura para a produção final de desenhos de cada especialidade; esta versão pode ser em arquivos eletrônicos, mas as modificações em relação à terceira etapa (segunda versão dos desenhos de trabalho de arquitetura) devem ser claramente marcadas, identificadas ou explicadas por escrito nos arquivos eletrônicos; os arquivos eletrônicos são gerenciados por camadas de acordo com as informações dos desenhos; após o recebimento, cada especialidade deve fazer os ajustes necessários de acordo com as características de expressão de seus próprios desenhos e excluir os gráficos, textos, cotas, etc. não relacionados à sua própria expressão",
        "ru": "Архитектурная специальность представляет текущую версию архитектурных чертежей и связанных материалов, которая должна отражать комплексные результаты после координации вышеуказанных процессов и служить архитектурной подосновой для окончательного выпуска чертежей каждой специальности; данная версия может быть в электронном виде, но изменения относительно третьего этапа (второй версии архитектурных рабочих чертежей) должны быть четко отмечены, обозначены или письменно разъяснены в электронных файлах; электронные файлы ведутся послойно в соответствии с информацией на чертежах; после получения каждая специальность должна внести необходимые корректировки в соответствии с особенностями отображения собственных чертежей и удалить графику, тексты, размеры и т.п., не относящиеся к собственному отображению",
        "ar": "يقدم تخصص العمارة النسخة الحالية من الرسومات المعمارية والمواد ذات الصلة، والتي يجب أن تعكس النتائج الشاملة بعد التنسيق بين العمليات المذكورة أعلاه، وتكون بمثابة الرسومات الأساسية المعمارية لإنتاج الرسومات النهائية لكل تخصص؛ يمكن أن تكون هذه النسخة في ملفات إلكترونية، لكن التعديلات مقارنة بالمرحلة الثالثة (النسخة الثانية من رسومات العمل المعمارية) يجب أن تكون محددة بوضوح أو مشار إليها أو موضحة كتابيًا في الملفات الإلكترونية؛ تُدار الملفات الإلكترونية في طبقات حسب المعلومات الموجودة في الرسومات؛ بعد الاستلام، يجب على كل تخصص إجراء التعديلات اللازمة وفقًا لخصائص التعبير لرسوماته الخاصة وحذف الأشكال والنصوص والأبعاد وغيرها غير المتعلقة بتعبير رسوماته الخاصة",
        "it": "La disciplina architettonica presenta la versione corrente degli elaborati architettonici e dei materiali correlati, che deve riflettere i risultati complessivi dopo il coordinamento dei processi sopra indicati e fungere da base architettonica per l'emissione finale degli elaborati di ciascuna disciplina; questa versione può essere in file elettronici, ma le modifiche rispetto alla terza fase (seconda versione degli elaborati operativi architettonici) devono essere chiaramente contrassegnate, identificate o spiegate per iscritto nei file elettronici; i file elettronici sono gestiti per livelli in base alle informazioni contenute negli elaborati; dopo la ricezione, ciascuna disciplina deve apportare le necessarie regolazioni in base alle caratteristiche espressive dei propri elaborati ed eliminare grafici, testi, quote, ecc. non pertinenti alla propria espressione",
        "vi": "Chuyên ngành kiến trúc trình bản vẽ kiến trúc bản hiện hành và tài liệu liên quan, phải phản ánh kết quả tổng hợp sau khi phối hợp các quá trình nêu trên, làm bản vẽ nền kiến trúc cho việc ra bản vẽ cuối cùng của từng chuyên ngành; bản vẽ bản này có thể ở dạng file điện tử, nhưng những chỗ sửa đổi so với giai đoạn thứ ba (bản vẽ tác nghiệp kiến trúc lần hai) phải được đánh dấu, nhận diện rõ ràng hoặc giải thích bằng văn bản trong file điện tử; file điện tử được quản lý theo lớp thông tin trong bản vẽ; sau khi nhận, mỗi chuyên ngành phải điều chỉnh cần thiết theo đặc điểm thể hiện bản vẽ của mình và xóa các hình, chữ, kích thước... không liên quan đến thể hiện bản vẽ của mình",
        "th": "สาขาสถาปัตยกรรมนำเสนอแบบสถาปัตยกรรมฉบับปัจจุบันและเอกสารที่เกี่ยวข้อง ซึ่งต้องสะท้อนผลลัพธ์รวมหลังจากการประสานงานของกระบวนการข้างต้น และใช้เป็นแบบฐานสถาปัตยกรรมสำหรับการผลิตแบบขั้นสุดท้ายของแต่ละสาขา ฉบับนี้อาจเป็นไฟล์อิเล็กทรอนิกส์ แต่จุดที่แก้ไขจากขั้นตอนที่สาม (แบบปฏิบัติงานสถาปัตยกรรมฉบับที่สอง) ต้องทำเครื่องหมาย ระบุ หรืออธิบายเป็นลายลักษณ์อักษรอย่างชัดเจนในไฟล์อิเล็กทรอนิกส์ ไฟล์อิเล็กทรอนิกส์ต้องจัดการเป็นชั้นตามข้อมูลในแบบ หลังจากรับแล้ว แต่ละสาขาต้องปรับปรุงตามลักษณะการแสดงแบบของตนและลบรูปทรง ตัวอักษร ขนาด ฯลฯ ที่ไม่เกี่ยวข้องกับการแสดงแบบของตน",
    },
    "在会签之前，各专业应再次进行专业间核对，确保上述各阶段的结论和成果已经得到完整落实；各专业的会审图纸应为已经过规定的校、审和修改程序后的图纸；会审完成后在各专业图纸上进行会签（设计文件归档之前完成）": {
        "en": "Before joint signing, each discipline shall re-verify among disciplines to ensure that the conclusions and results of the above stages have been fully implemented; the drawings for joint review of each discipline shall have passed the prescribed procedures of checking, reviewing and revising; after the joint review is completed, joint signing shall be made on the drawings of each discipline (to be completed before the design documents are archived)",
        "ja": "会签前に、各専門は再び専門間の照合を行い、上記各段階の結論と成果が完全に実施されていることを確保する；各専門の会審図面は規定の校・審・修正手順を経た図面でなければならない；会審完了後、各専門図面に会签を行う（設計文書のアーカイブ前に完了）",
        "ko": "회서(會簽) 전에 각 전문은 다시 전문 간 대조를 수행하여 상기 각 단계의 결론과 성과가 완전히 이행되었음을 보장한다；각 전문의 합동 심사 도면은 규정된 검교·심사·수정 절차를 거친 도면이어야 한다；합동 심사 완료 후 각 전문 도면에 회서를 진행한다(설계 문서 보관 전 완료)",
        "fr": "Avant la signature conjointe, chaque spécialité doit procéder à nouveau à une vérification interdisciplinaire afin de garantir que les conclusions et les résultats des étapes ci-dessus sont pleinement mis en œuvre ; les plans de révision conjointe de chaque spécialité doivent avoir passé les procédures prescrites de vérification, de révision et de modification ; une fois la révision conjointe achevée, la signature conjointe est apposée sur les plans de chaque spécialité (à terminer avant l'archivage des documents de conception)",
        "de": "Vor der gemeinsamen Unterzeichnung muss jede Fachrichtung erneut fachübergreifende Abstimmungen vornehmen, um sicherzustellen, dass die Schlussfolgerungen und Ergebnisse der oben genannten Phasen vollständig umgesetzt wurden; die Pläne zur gemeinsamen Prüfung jeder Fachrichtung müssen die vorgeschriebenen Prüfungs-, Revisions- und Änderungsverfahren durchlaufen haben; nach Abschluss der gemeinsamen Prüfung ist auf den Plänen jeder Fachrichtung zu unterzeichnen (vor der Archivierung der Entwurfsunterlagen abzuschließen)",
        "es": "Antes de la firma conjunta, cada especialidad debe volver a realizar la verificación interdisciplinaria para garantizar que las conclusiones y resultados de las etapas anteriores se han implementado plenamente; los planos de revisión conjunta de cada especialidad deben haber pasado los procedimientos prescritos de verificación, revisión y modificación; una vez completada la revisión conjunta, se procederá a la firma conjunta en los planos de cada especialidad (debe completarse antes del archivo de los documentos de diseño)",
        "pt": "Antes da assinatura conjunta, cada especialidade deve realizar novamente a verificação interdisciplinar para garantir que as conclusões e resultados das etapas acima foram plenamente implementados; os desenhos para revisão conjunta de cada especialidade devem ter passado pelos procedimentos prescritos de verificação, revisão e modificação; após a conclusão da revisão conjunta, procede-se à assinatura conjunta nos desenhos de cada especialidade (a concluir antes do arquivamento dos documentos de projeto)",
        "ru": "Перед совместным подписанием каждая специальность должна повторно провести межспециальную сверку, чтобы убедиться, что выводы и результаты вышеуказанных этапов полностью реализованы; чертежи для совместной экспертизы каждой специальности должны пройти установленные процедуры проверки, рецензирования и внесения изменений; после завершения совместной экспертизы на чертежах каждой специальности производится совместное подписание (завершить до сдачи проектной документации в архив)",
        "ar": "قبل التوقيع المشترك، يجب على كل تخصص إجراء مراجعة مشتركة بين التخصصات مرة أخرى للتأكد من أن استنتاجات ونتائج المراحل المذكورة أعلاه قد نُفذت بالكامل؛ يجب أن تكون رسومات المراجعة المشتركة لكل تخصص قد خضعت لإجراءات الفحص والمراجعة والتعديل المقررة؛ بعد اكتمال المراجعة المشتركة، يتم التوقيع المشترك على رسومات كل تخصص (يُكمل قبل أرشفة مستندات التصميم)",
        "it": "Prima della firma congiunta, ogni disciplina deve effettuare nuovamente la verifica interdisciplinare per garantire che le conclusioni e i risultati delle fasi sopra indicate siano stati pienamente attuati; gli elaborati per la revisione congiunta di ciascuna disciplina devono aver superato le procedure prescritte di verifica, revisione e modifica; completata la revisione congiunta, si procede alla firma congiunta sugli elaborati di ciascuna disciplina (da completare prima dell'archiviazione dei documenti di progetto)",
        "vi": "Trước khi ký chung, mỗi chuyên ngành phải tiến hành đối chiếu liên chuyên ngành lần nữa để bảo đảm các kết luận và kết quả của các giai đoạn nêu trên đã được thực hiện đầy đủ; bản vẽ hội thẩm của mỗi chuyên ngành phải đã qua các quy trình kiểm tra, thẩm tra và sửa đổi quy định; sau khi hoàn tất hội thẩm, tiến hành ký chung trên bản vẽ của từng chuyên ngành (hoàn thành trước khi lưu trữ hồ sơ thiết kế)",
        "th": "ก่อนการลงนามร่วมกัน แต่ละสาขาต้องดำเนินการตรวจสอบระหว่างสาขาอีกครั้งเพื่อให้แน่ใจว่าข้อสรุปและผลลัพธ์ของขั้นตอนข้างต้นได้ถูกนำไปปฏิบัติอย่างครบถ้วน แบบสำหรับการตรวจร่วมของแต่ละสาขาต้องผ่านขั้นตอนการตรวจทาน ตรวจสอบ และแก้ไขตามที่กำหนดไว้ เมื่อการตรวจร่วมเสร็จสิ้น ให้ลงนามร่วมกันบนแบบของแต่ละสาขา (แล้วเสร็จก่อนการจัดเก็บเอกสารการออกแบบ)",
    },
    "电提暖：特殊房间温湿度参数和通风要求、电气发热量较大的房间设备发热量（或估算电气安装容量，如变配电室、大型计算机主机房）": {"en": "electrical to HVAC: temperature and humidity parameters and ventilation requirements of special rooms, heat gain of equipment in rooms with large electrical heat release (or estimated electrical installation capacity, such as substations and large computer mainframe rooms)", "ja": "電→暖：特殊室の温湿度パラメータと換気要求、電気発熱量が大きい室の機器発熱量（または概算電気設備容量。変配電室、大型コンピュータ室など）", "ko": "전기→난방: 특수실 온습도 파라미터와 환기 요구, 전기 발열량이 큰 실의 설비 발열량(또는 추정 전기 설비 용량, 예: 변배전실, 대형 컴퓨터 주기기실)", "fr": "électricité vers CVC : paramètres de température et d'humidité et exigences de ventilation des locaux spéciaux, dégagement thermique des équipements des locaux à forte charge électrique (ou capacité électrique installée estimée, comme les postes de transformation et les grandes salles d'ordinateurs centraux)", "de": "Elektro an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter sowie Lüftungsanforderungen besonderer Räume, Wärmeabgabe von Geräten in Räumen mit hoher elektrischer Wärmeentwicklung (oder geschätzte elektrische Anschlussleistung, z. B. Umspannwerke und große Rechnerräume)", "es": "electricidad a climatización: parámetros de temperatura y humedad y requisitos de ventilación de locales especiales, carga térmica de los equipos de locales con gran disipación eléctrica (o capacidad eléctrica instalada estimada, como subestaciones y grandes salas de computadoras)", "pt": "elétrica para climatização: parâmetros de temperatura e umidade e requisitos de ventilação de ambientes especiais, carga térmica dos equipamentos de ambientes com grande dissipação elétrica (ou capacidade elétrica instalada estimada, como subestações e grandes salas de computadores)", "ru": "электроснабжение — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции специальных помещений, тепловыделение оборудования в помещениях с большим тепловыделением от электрооборудования (или расчётная установленная мощность, например трансформаторные подстанции, крупные машинные залы)", "ar": "والكهرباء إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية للغرف الخاصة، والحرارة المنبعثة من معدات الغرف ذات الإطلاق الكهربائي الكبير (أو السعة الكهربائية المركبة التقديرية، مثل محطات التحويل وغرف الحواسيب الرئيسية الكبيرة)", "it": "impianti elettrici a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali speciali, carico termico delle apparecchiature dei locali con grande dissipazione elettrica (o potenza elettrica installata stimata, come cabine elettriche e grandi sale computer)", "vi": "điện gửi điều hòa: thông số nhiệt độ độ ẩm và yêu cầu thông gió của phòng đặc biệt, lượng nhiệt tỏa ra của thiết bị trong phòng có lượng tỏa nhiệt điện lớn (hoặc công suất lắp đặt điện ước tính, như trạm biến áp, phòng máy tính lớn)", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่มีการทำความร้อน ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ปริมาณความร้อนจากอุปกรณ์ในห้องที่มีการปลดปล่อยความร้อนทางไฟฟ้าสูง (หรือกำลังไฟฟ้าติดตั้งโดยประมาณ เช่น ห้องสถานีย่อย ห้องคอมพิวเตอร์เมนเฟรมขนาดใหญ่) ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องอุปกรณ์ไฟฟ้าอ่อนและห้องควบคุมระบบรักษาความปลอดภัย ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณการใช้ไฟฟ้าของห้องอุปกรณ์ไฟฟ้าอ่อน ห้องควบคุมระบบรักษาความปลอดภัย และปล่องไฟฟ้าอ่อน"},
    "电提建筑：补充完善确认变配电所、发电机房等机房面积尺寸净高，各层弱电机房配电小室位置面积，主要强电竖井位置尺寸，楼板厚度要求，发电机冷却进风与排风竖井位置面积": {"en": "electrical to architectural: supplement, improve and confirm the area, dimensions and clear height of substations, generator rooms and other plant rooms, the positions and areas of weak-current rooms and distribution cubicles on each floor, the position and dimensions of main power riser shafts, the floor slab thickness requirement, and the positions and areas of generator cooling air intake and exhaust shafts", "ja": "電提建築：変電所・発電機室等の機械室面積・寸法・天井高を補完・改善・確認し、各階の弱電室・配電小室の位置面積、主要強電シャフトの位置寸法、床スラブ厚さ要求、発電機冷却の給気・排気シャフトの位置面積を提出する", "ko": "전기가 건축에: 변전소·발전기실 등 기계실의 면적·치수·천장고를 보완·개선·확인하고, 각 층 약전실·배전 소실의 위치 면적, 주요 강전 샤프트 위치 치수, 슬래브 두께 요구, 발전기 냉각 흡기·배기 샤프트 위치 면적을 제출합니다", "fr": "électricité à architecture : compléter, améliorer et confirmer la surface, les dimensions et la hauteur libre des postes de transformation, groupes électrogènes et autres locaux techniques, les positions et surfaces des locaux courants faibles et des petites armoires de distribution à chaque niveau, les positions et dimensions des gaines électriques principales, l'exigence d'épaisseur des dalles, les positions et surfaces des conduits d'aspiration et d'évacuation du refroidissement des groupes électrogènes", "de": "Elektro an Architektur: Ergänzung, Verbesserung und Bestätigung der Fläche, Abmessungen und lichten Höhe von Trafostationen, Generatorräumen usw., der Positionen und Flächen von Schwachstromräumen und Verteilerkabinen je Geschoss, der Position und Abmessungen der Hauptstromschächte, der Anforderung an die Geschossdeckenstärke und der Positionen und Flächen von Zuluft- und Abluftschächten der Generatorkühlung", "es": "electricidad a arquitectura: completar, mejorar y confirmar la superficie, las dimensiones y la altura libre de las subestaciones, salas de grupos electrógenos y otros locales técnicos, las posiciones y superficies de los locales de corrientes débiles y pequeños cuartos de distribución en cada nivel, las posiciones y dimensiones de los patinillos eléctricos principales, el requisito de espesor de los forjados, las posiciones y superficies de los conductos de admisión y extracción del enfriamiento de los grupos electrógenos", "pt": "elétrica para arquitetura: complementar, melhorar e confirmar a área, as dimensões e o pé-direito de subestações, salas de geradores e outras salas técnicas, as posições e áreas dos ambientes de corrente fraca e pequenas salas de distribuição em cada pavimento, as posições e dimensões dos shaft elétricos principais, a exigência de espessura das lajes, as posições e áreas dos dutos de admissão e exaustão do resfriamento dos geradores", "ru": "электроснабжение — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты трансформаторных подстанций, дизельных и т.п., расположение и площади слаботочных помещений и распределительных ниш на каждом этаже, положение и размеры основных силовых шахт, требование к толщине плит перекрытий, расположение и площади шахт приточной и вытяжной вентиляции охлаждения генераторов", "ar": "والكهرباء إلى العمارة: استكمال وتحسين وتأكيد مساحة وأبعاد وارتفاع المحطات الفرعية وغرف المولدات وغيرها من الغرف الفنية، ومواقع ومساحات غرف التيارات الضعيفة وحجرات التوزيع الصغيرة في كل طابق، ومواقع وأبعاد أعمدة التيار القوي الرئيسية، ومتطلب سمك الألواح، ومواقع ومساحات أعمدة سحب وطرد هواء تبريد المولدات", "it": "elettrici ad architettura: integrare, migliorare e confermare l'area, le dimensioni e l'altezza libera di cabine di trasformazione, sale generatori e altri locali tecnici, le posizioni e aree dei locali debole corrente e dei piccoli locali di distribuzione a ogni piano, le posizioni e dimensioni dei vani elettrici principali, il requisito di spessore delle solette, le posizioni e aree dei condotti di aspirazione e scarico del raffreddamento dei generatori", "vi": "điện gửi kiến trúc: bổ sung hoàn thiện xác nhận diện tích, kích thước, chiều cao thông thủy trạm biến áp, phòng máy phát, v.v., vị trí diện tích phòng điện yếu và phòng phân phối nhỏ từng tầng, vị trí kích thước giếng điện mạnh chính, yêu cầu chiều dày sàn, vị trí diện tích giếng hút/ thải gió làm mát máy phát", "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น"},
    "智能化提建筑：补充完善确认进线间、网络机房、电话交换机房、有线电视机房、手机信号覆盖机房、安防控制室、卫星电视等机房面积尺寸净高，弱电竖井大型会议室控制室位置尺寸，楼板厚度要求": {"en": "intelligent systems to architectural: supplement, improve and confirm the area, dimensions and clear height of incoming line rooms, network rooms, telephone exchange rooms, cable TV rooms, mobile signal coverage rooms, security control rooms, satellite TV rooms, etc., the position and dimensions of weak-current riser shafts, large meeting room control rooms, and the floor slab thickness requirement", "ja": "智能化提建築：引込線室・ネットワーク室・電話交換室・有線テレビ室・携帯信号カバー室・安防控制室・衛星テレビ等の機械室面積・寸法・天井高を補完・改善・確認し、弱電シャフト・大型会議室制御室の位置寸法、床スラブ厚さ要求を提出する", "ko": "지능화가 건축에: 인입선실·네트워크실·전화교환실·유선TV실·휴대폰 신호 커버실·안보 통제실·위성TV 등 기계실의 면적·치수·천장고를 보완·개선·확인하고, 약전 샤프트·대형 회의실 제어실의 위치 치수, 슬래브 두께 요구를 제출합니다", "fr": "systèmes intelligents à architecture : compléter, améliorer et confirmer la surface, les dimensions et la hauteur libre des locaux de pénétration, locaux réseau, locaux de centraux téléphoniques, locaux de télévision par câble, locaux de couverture de signaux mobiles, salles de contrôle de sécurité, locaux de télévision par satellite, etc., les positions et dimensions des gaines courants faibles et des salles de contrôle de grandes salles de réunion, et l'exigence d'épaisseur des dalles", "de": "Gebäudeautomation an Architektur: Ergänzung, Verbesserung und Bestätigung der Fläche, Abmessungen und lichten Höhe von Einführungsräumen, Netzwerkräumen, Telefonvermittlungsräumen, Kabelfernsehräumen, Mobilfunksignale-Versorgungsräumen, Sicherheitskontrollräumen, Satellitenfernsehräumen usw., der Position und Abmessungen von Schwachstromschächten, Kontrollräumen großer Konferenzräume und der Anforderung an die Geschossdeckenstärke", "es": "sistemas inteligentes a arquitectura: completar, mejorar y confirmar la superficie, las dimensiones y la altura libre de los locales de acometida, locales de red, locales de centralitas telefónicas, locales de televisión por cable, locales de cobertura de señal móvil, salas de control de seguridad, locales de televisión por satélite, etc., las posiciones y dimensiones de los patinillos de corrientes débiles y de las salas de control de grandes salas de reuniones, y el requisito de espesor de los forjados", "pt": "sistemas inteligentes para arquitetura: complementar, melhorar e confirmar a área, as dimensões e o pé-direito de salas de entrada, salas de rede, salas de centrais telefônicas, salas de TV a cabo, salas de cobertura de sinal móvel, salas de controle de segurança, salas de TV via satélite, etc., as posições e dimensões dos shaft de corrente fraca e das salas de controle de grandes salas de reunião, e a exigência de espessura das lajes", "ru": "интеллектуальные системы — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений ввода, серверных, АТС, кабельного телевидения, покрытия сотовой связи, охранных диспетчерских, спутникового телевидения и т.п., положение и размеры слаботочных шахт, диспетчерских больших конференц-залов и требование к толщине плит перекрытий", "ar": "والأنظمة الذكية إلى العمارة: استكمال وتحسين وتأكيد مساحة وأبعاد وارتفاع غرف خطوط الدخول وغرف الشبكات وغرف مقاسم الهاتف وغرف التلفزيون الكبلي وغرف تغطية إشارات الهاتف المحمول وغرف التحكم الأمنية وغرف التلفزيون الفضائي وغيرها، ومواقع وأبعاد أعمدة التيارات الضعيفة وغرف التحكم في قاعات الاجتماعات الكبيرة، ومتطلب سمك الألواح", "it": "sistemi intelligenti ad architettura: integrare, migliorare e confermare l'area, le dimensioni e l'altezza libera di locali di ingresso linea, locali rete, locali centralini telefonici, locali TV via cavo, locali di copertura segnale mobile, sale controllo sicurezza, locali TV via satellite, ecc., le posizioni e dimensioni dei vani debole corrente e delle sale di controllo delle grandi sale riunioni, e il requisito di spessore delle solette", "vi": "hệ thống thông minh gửi kiến trúc: bổ sung hoàn thiện xác nhận diện tích, kích thước, chiều cao thông thủy phòng cáp vào, phòng mạng, phòng tổng đài điện thoại, phòng truyền hình cáp, phòng phủ sóng di động, phòng điều khiển an ninh, phòng truyền hình vệ tinh, v.v., vị trí kích thước giếng điện yếu, phòng điều khiển phòng họp lớn, yêu cầu chiều dày sàn", "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น"},
    "智能化提电：各弱电机房、安防控制室、弱电间用电量": {"en": "intelligent systems to electrical: power consumption of weak-current equipment rooms, security control rooms and weak-current rooms", "ja": "智能化→電：各弱電機械室・安防控制室・弱電間の電力消費量", "ko": "지능화→전기: 각 약전 기계실·보안 통제실·약전실의 전력 소비량", "fr": "systèmes intelligents vers électricité : consommation électrique des locaux techniques courants faibles, des salles de contrôle de sécurité et des gaines courants faibles", "de": "Gebäudeautomation an Elektro: Stromverbrauch von Schwachstrom-Räumen, Sicherheitskontrollräumen und Schwachstrom-Schächten", "es": "sistemas inteligentes a electricidad: consumo eléctrico de los cuartos de corrientes débiles, salas de control de seguridad y patinillos de corrientes débiles", "pt": "sistemas inteligentes para elétrica: consumo elétrico das salas de corrente fraca, salas de controle de segurança e shaft de corrente fraca", "ru": "интеллектуальные системы — электроснабжению: потребляемая мощность слаботочных аппаратных, помещений охраны и слаботочных помещений", "ar": "والأنظمة الذكية إلى الكهرباء: استهلاك الكهرباء لغرف التيارات الضعيفة وغرف التحكم الأمنية وأعمدة التيارات الضعيفة", "it": "sistemi intelligenti a elettrici: consumo elettrico dei locali debole corrente, delle sale controllo sicurezza e dei vani debole corrente", "vi": "hệ thống thông minh gửi điện: lượng điện tiêu thụ của các phòng máy yếu điện, phòng điều khiển an ninh và phòng yếu điện", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่มีการทำความร้อน ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ปริมาณความร้อนจากอุปกรณ์ในห้องที่มีการปลดปล่อยความร้อนทางไฟฟ้าสูง (หรือกำลังไฟฟ้าติดตั้งโดยประมาณ เช่น ห้องสถานีย่อย ห้องคอมพิวเตอร์เมนเฟรมขนาดใหญ่) ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องอุปกรณ์ไฟฟ้าอ่อนและห้องควบคุมระบบรักษาความปลอดภัย ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณการใช้ไฟฟ้าของห้องอุปกรณ์ไฟฟ้าอ่อน ห้องควบคุมระบบรักษาความปลอดภัย และปล่องไฟฟ้าอ่อน"},
    "水、暖、电、智能化提结构：吊挂管道和设备的楼板荷载；需要结构专业进行设计的大型设备的基础尺寸、设备重量及相关要求；水、暖专业中需要结构专业进行加固的管道固定支架的位置和推力": {"en": "Plumbing, HVAC, electrical and intelligent systems to structural: the floor slab loads of suspended pipes and equipment；the foundation dimensions, equipment weights and related requirements of large equipment requiring structural design；the positions and thrusts of pipe fixed supports in the plumbing and HVAC disciplines that require structural reinforcement", "ja": "水・暖・電・智能化提構造：吊下げ配管と設備の床版荷重；構造専門の設計が必要な大型設備の基礎寸法、設備重量及び関連要求；水・暖専門の中で構造専門による補強が必要な配管固定支持架の位置と推力", "ko": "급배수·난방·전기·지능화가 구조에: 매달린 배관과 설비의 바닥 슬래브 하중；구조 전문의 설계가 필요한 대형 설비의 기초 치수, 설비 중량 및 관련 요구；급배수·난방 전문 중 구조 전문의 보강이 필요한 배관 고정 지지대의 위치와 추력", "fr": "Plomberie, CVC, électricité et systèmes intelligents à structure : les charges sur dalles des conduits suspendus et des équipements；les dimensions des fondations, les poids des équipements et les exigences connexes des grands équipements nécessitant une conception structurelle；les positions et poussées des supports fixes de canalisations dans les disciplines plomberie et CVC nécessitant un renforcement structurel", "de": "Sanitär, Heizung-Lüftung, Elektro und Gebäudeautomation an Tragwerk: die Deckenlasten hängender Rohrleitungen und Anlagen；die Gründungsabmessungen, Anlagengewichte und damit verbundenen Anforderungen großer Anlagen, die eine statische Bemessung erfordern；die Positionen und Schubkräfte von Rohrleitungsfestpunkten in den Disziplinen Sanitär und Heizung-Lüftung, die eine bauliche Verstärkung erfordern", "es": "Fontanería, climatización, electricidad y sistemas inteligentes a estructuras: las cargas sobre forjados de los conductos colgados y equipos；las dimensiones de cimentación, pesos de los equipos y requisitos relacionados de los grandes equipos que requieren diseño estructural；las posiciones y empujes de los soportes fijos de tuberías en las especialidades de fontanería y climatización que requieren refuerzo estructural", "pt": "Hidráulica, climatização, elétrica e sistemas inteligentes para estruturas: as cargas sobre lajes de tubulações suspensas e equipamentos；as dimensões de fundações, pesos dos equipamentos e requisitos relacionados dos grandes equipamentos que exigem projeto estrutural；as posições e empuxos dos suportes fixos de tubulações nas especialidades de hidráulica e climatização que exigem reforço estrutural", "ru": "Водоснабжение, отопление и вентиляция, электроснабжение и интеллектуальные системы — конструкциям: нагрузки на перекрытия от подвесных трубопроводов и оборудования；размеры фундаментов, массы оборудования и связанные требования для крупного оборудования, требующего расчёта конструкций；расположение и распор неподвижных опор трубопроводов в специальностях водоснабжения и отопления и вентиляции, требующих усиления конструкций", "ar": "السباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية إلى الإنشاءات: الأحمال على الألواح من الأنابيب المعلقة والمعدات；وأبعاد الأساسات وأوزان المعدات والمتطلبات ذات الصلة للمعدات الكبيرة التي تتطلب تصميمًا إنشائياً；ومواقع وقوى الدفع للدعامات الثابتة للأنابيب في تخصصي السباكة والتدفئة والتهوية التي تتطلب تقوية إنشائية", "it": "Idrici, HVAC, elettrici e sistemi intelligenti a strutture: i carichi sui solai di tubazioni sospese e apparecchiature；le dimensioni delle fondazioni, i pesi delle apparecchiature e i requisiti correlati delle grandi apparecchiature che richiedono progettazione strutturale；le posizioni e le spinte dei supporti fissi delle tubazioni nelle discipline idriche e HVAC che richiedono rinforzo strutturale", "vi": "Cấp thoát nước, điều hòa, điện, hệ thống thông minh gửi kết cấu: tải trọng sàn của đường ống treo và thiết bị；kích thước móng, trọng lượng thiết bị và yêu cầu liên quan của thiết bị lớn cần kết cấu thiết kế；vị trí và lực đẩy của gối đỡ cố định đường ống trong chuyên ngành cấp thoát nước, điều hòa cần kết cấu gia cường", "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: น้ำหนักบรรทุกบนพื้นจากท่อแขวนและอุปกรณ์ ขนาดฐานราก น้ำหนักอุปกรณ์ และข้อกำหนดที่เกี่ยวข้องของอุปกรณ์ขนาดใหญ่ที่ต้องออกแบบโครงสร้าง ตำแหน่งและแรงดันของขายึดท่อคงที่ในสาขาประปาและระบบปรับอากาศที่ต้องเสริมโครงสร้าง ประปาและระบบปรับอากาศส่งให้สถาปัตยกรรม: ข้อกำหนดตำแหน่ง ขนาด และความลาดเอียงของรางระบายน้ำ ประปาส่งให้สถาปัตยกรรม: ข้อกำหนดของห้องพิเศษ (เช่น ห้องดับเพลิงด้วยแก๊ส ฯลฯ) ต่อช่องเปิด เช่น ประตูและหน้าต่าง"},
    "水、暖提电：各强电设备的位置、电压等级及其电气安装容量（消防设备和平时使用的设备分别提出）": {"en": "plumbing and HVAC to electrical: the positions, voltage levels and electrical installed capacities of various strong-current equipment (fire protection equipment and equipment for normal use proposed separately)", "ja": "水・暖提電：各強電設備の位置、電圧レベル及びその電気設備容量（消防設備と平時使用設備は別々に提出）", "ko": "급배수·난방이 전기에: 각 강전 설비의 위치, 전압 레벨 및 전기 설비 용량(소방 설비와 평시 사용 설비는 별도 제출)", "fr": "plomberie et CVC à électricité : les positions, niveaux de tension et puissances installées électriques des divers équipements courants forts (les équipements de protection incendie et ceux d'usage courant proposés séparément)", "de": "Sanitär und Heizung-Lüftung an Elektro: die Positionen, Spannungsebenen und elektrischen Anschlussleistungen der verschiedenen Starkstromanlagen (Brandschutzanlagen und Anlagen für den Normalbetrieb getrennt angeben)", "es": "fontanería y climatización a electricidad: las posiciones, niveles de tensión y potencias eléctricas instaladas de los diversos equipos de corrientes fuertes (los equipos de protección contra incendios y los de uso normal se proponen por separado)", "pt": "hidráulica e climatização para elétrica: as posições, níveis de tensão e potências instaladas elétricas dos diversos equipamentos de corrente forte (equipamentos de proteção contra incêndio e de uso normal propostos separadamente)", "ru": "водоснабжение и отопление и вентиляция — электроснабжению: расположение, уровни напряжения и установленные электрические мощности различных сильноточных устройств (противопожарное оборудование и оборудование повседневного использования указываются отдельно)", "ar": "والسباكة والتدفئة والتهوية إلى الكهرباء: مواقع ومستويات الجهد والقدرات الكهربائية المركبة لمختلف معدات التيار القوي (يتم اقتراح معدات الحماية من الحرائق ومعدات الاستخدام العادي بشكل منفصل)", "it": "idrici e HVAC a elettrici: le posizioni, i livelli di tensione e le potenze elettriche installate delle varie apparecchiature di corrente forte (le apparecchiature antincendio e quelle di uso ordinario proposte separatamente)", "vi": "cấp thoát nước, điều hòa gửi điện: vị trí, cấp điện áp và công suất lắp đặt điện của các thiết bị điện mạnh (thiết bị phòng cháy và thiết bị sử dụng thường xuyên nêu riêng)", "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: น้ำหนักบรรทุกโครงสร้างของอุปกรณ์ห้องเครื่องและหอหล่อเย็นที่อยู่บนพื้น ตำแหน่งเบื้องต้นของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนักและพื้น ประปาและระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่ง ระดับแรงดันไฟฟ้า และกำลังไฟฟ้าติดตั้งของอุปกรณ์ไฟฟ้าแรงต่าง ๆ (อุปกรณ์ดับเพลิงและอุปกรณ์ใช้ปกติเสนอแยกกัน)"},
    "水提总图：各市政管网（给水、排水、热水等）至本楼接口大致位置，主要室外管道布置及主要构筑物（化粪池、隔油池、水表井、水泵接合器等）大致位置": {"en": "Plumbing to site planning: approximate positions of the interfaces of various municipal pipeline networks (water supply, drainage, hot water, etc.) to this building, approximate positions of major outdoor pipeline layouts and main structures (septic tanks, grease traps, water meter wells, fire pump connections, etc.)", "ja": "水提総図：各市政管網（給水、排水、給湯など）から本棟への接続おおよその位置、主要屋外配管配置及び主要構築物（化糞池、グリーストラップ、水道メーター井、ポンプ接合器など）のおおよその位置", "ko": "급배수가 총도에: 각 시정 관망(급수, 배수, 온수 등)의 본 건물 접속 대략 위치, 주요 옥외 배관 배치 및 주요 구조물(정화조, 유분리조, 수도계량기 우물, 펌프 접합기 등) 대략 위치", "fr": "Plomberie à l'aménagement du site : positions approximatives des interfaces des divers réseaux municipaux (adduction, assainissement, eau chaude, etc.) avec ce bâtiment, positions approximatives des principaux réseaux extérieurs et ouvrages (fosses septiques, séparateurs à graisses, regards de compteurs d'eau, raccordements de pompes d'incendie, etc.)", "de": "Sanitär an Standortplanung: ungefähre Positionen der Anschlusspunkte der verschiedenen städtischen Netze (Wasserversorgung, Entwässerung, Warmwasser usw.) an dieses Gebäude, ungefähre Positionen der Hauptaußenleitungen und der wichtigsten Bauwerke (Kläranlagen, Fettabscheider, Wasserzählerschächte, Feuerlöschpumpenanschlüsse usw.)", "es": "Fontanería a urbanización: posiciones aproximadas de los puntos de conexión de las distintas redes municipales (abastecimiento, saneamiento, agua caliente, etc.) con este edificio, posiciones aproximadas de las principales redes exteriores y obras (fosas sépticas, separadores de grasas, arquetas de contadores de agua, conexiones de bombas de incendio, etc.)", "pt": "Hidráulica para urbanismo: posições aproximadas dos pontos de conexão das diversas redes municipais (abastecimento, drenagem, água quente, etc.) com este edifício, posições aproximadas dos principais traçados externos e obras (fossas sépticas, separadores de gordura, caixas de hidrômetros, conexões de bombas de incêndio, etc.)", "ru": "Водоснабжение — генплану: приблизительное расположение точек присоединения городских сетей (водоснабжение, водоотведение, горячее водоснабжение и т.п.) к данному зданию, приблизительное расположение основных наружных сетей и сооружений (септики, жироуловители, колодцы водомеров, пожарные соединительные головки и т.п.)", "ar": "السباكة إلى تخطيط المواقع: المواقع التقريبية لنقاط ربط شبكات المرافق البلدية المختلفة (إمداد المياه، الصرف، الماء الساخن وغيرها) بهذا المبنى، والمواقع التقريبية للشبكات الخارجية الرئيسية والمنشآت الرئيسية (خزانات الصرف الصحي، فواصل الدهون، آبار عدادات المياه، وصلات مضخات الإطفاء وغيرها)", "it": "Idrici a pianificazione del sito: posizioni approssimative dei punti di collegamento delle varie reti comunali (acquedotto, fognatura, acqua calda, ecc.) con questo edificio, posizioni approssimative dei principali tracciati esterni e delle opere (fosse settiche, separatori di grassi, pozzetti contatori acqua, attacchi pompe antincendio, ecc.)", "vi": "Cấp thoát nước gửi tổng mặt bằng: vị trí tương đối các điểm đấu nối của các mạng hạ tầng (cấp nước, thoát nước, nước nóng, v.v.) vào công trình này, vị trí tương đối bố trí đường ống ngoài nhà chính và công trình chính (bể tự hoại, bể tách dầu mỡ, giếng đồng hồ nước, đầu nối máy bơm chữa cháy, v.v.)", "th": "ประปาส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อเครือข่ายสาธารณูปโภคต่าง ๆ (ประปา ท่อระบายน้ำ น้ำร้อน ฯลฯ) เข้าสู่อาคารนี้ ตำแหน่งโดยประมาณของการวางท่อภายนอกหลักและโครงสร้างหลัก (ถังบำบัด ถังดักไขมัน บ่อมาตรวัดน้ำ ข้อต่อปั๊มดับเพลิง ฯลฯ) ระบบปรับอากาศส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อเครือข่ายสาธารณูปโภค (ความร้อน แก๊ส) เข้าสู่อาคาร ตำแหน่งโดยประมาณของการวางท่อภายนอกหลักและโครงสร้าง (บ่อวาล์ว ฯลฯ) ไฟฟ้าและระบบอัจฉริยะส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อสายสาธารณูปโภค (ไฟฟ้า การสื่อสาร เคเบิลทีวี ฯลฯ) เข้าสู่อาคาร ตำแหน่งโดยประมาณของการวางสายภายนอกหลักและโครงสร้าง (บ่อตรวจสอบ ฯลฯ)"},
    "暖提建筑：补齐全部管道井尺寸和位置要求、补齐全部室外立面百页窗（新风口、排风口等）位置和尺寸": {"en": "HVAC to architectural: complete the size and position requirements of all pipe shafts, and complete the positions and dimensions of all exterior louver windows (fresh air outlets, exhaust outlets, etc.)", "ja": "暖提建築：すべてのパイプシャフト寸法と位置要求、すべての室外立面ルーバー窓（新風口、排風口など）の位置と寸法を補完", "ko": "난방이 건축에: 모든 파이프 샤프트 치수와 위치 요구, 모든 외부 입면 루버 창(신풍구, 배풍구 등)의 위치와 치수를 보완", "fr": "CVC à architecture : compléter les exigences de dimensions et de positions de toutes les gaines techniques et compléter les positions et dimensions de toutes les fenêtres à persiennes en façade (bouches d'air neuf, bouches d'extraction, etc.)", "de": "Heizung-Lüftung an Architektur: die Größen- und Positionsanforderungen aller Installationsschächte sowie die Positionen und Abmessungen aller außenliegenden Lamellenfenster (Frischluftöffnungen, Abluftöffnungen usw.) vervollständigen", "es": "climatización a arquitectura: completar los requisitos de dimensiones y posiciones de todos los patinillos y completar las posiciones y dimensiones de todas las ventanas de lamas de fachada (tomas de aire nuevo, salidas de extracción, etc.)", "pt": "climatização para arquitetura: completar os requisitos de dimensões e posições de todos os shafts e completar as posições e dimensões de todas as janelas de lâminas na fachada (entradas de ar novo, saídas de exaustão, etc.)", "ru": "отопление и вентиляция — архитектуре: дополнить требования к размерам и расположению всех инженерных шахт и дополнить расположение и размеры всех фасадных жалюзийных решёток (воздухозаборов, вытяжных отверстий и т.п.)", "ar": "والتدفئة والتهوية إلى العمارة: استكمال متطلبات الأبعاد والمواقع لجميع الأعمدة الفنية، واستكمال مواقع وأبعاد جميع نوافذ الشرائح الخارجية (مداخل الهواء الجديد، مخارج الهواء العادم وغيرها)", "it": "HVAC a architettura: completare i requisiti di dimensioni e posizioni di tutti i cavedi e completare le posizioni e dimensioni di tutte le finestre a lamelle in facciata (prese d'aria nuova, bocchette di estrazione, ecc.)", "vi": "điều hòa gửi kiến trúc: bổ sung đầy đủ yêu cầu kích thước và vị trí các hộp kỹ thuật đường ống, bổ sung đầy đủ vị trí và kích thước cửa chớp mặt đứng (miệng gió tươi, miệng gió thải, v.v.)", "th": "ประปาส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปาส่งให้ไฟฟ้า: ตำแหน่งหัวจ่ายน้ำดับเพลิง เครื่องแสดงการไหลของน้ำ วาล์วสัญญาณเตือน วาล์วควบคุม ฯลฯ และข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ระบบปรับอากาศส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด และเติมเต็มตำแหน่งและขนาดของช่องระบายอากาศแบบใบพัดภายนอกทั้งหมด (ช่องลมเข้า ช่องลมออก ฯลฯ) ระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งวาล์วระบายควัน วาล์วส่งลมอัดแรงดัน วาล์วกันไฟที่ต้องควบคุมด้วยไฟฟ้า และวาล์วควบคุมอัตโนมัติ รวมถึงข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ไฟฟ้าส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งที่แม่นยำของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนัก พื้น และคาน"},
    "水、暖提建筑：排水沟的位置、尺寸、坡度等要求": {"en": "plumbing and HVAC to architectural: the requirements for the positions, dimensions and slopes of drainage ditches", "ja": "水・暖提建築：排水溝の位置、寸法、勾配などの要求", "ko": "급배수·난방이 건축에: 배수구의 위치, 치수, 기울기 등의 요구", "fr": "plomberie et CVC à architecture : les exigences de positions, dimensions et pentes des caniveaux d'évacuation", "de": "Sanitär und Heizung-Lüftung an Architektur: die Anforderungen an Position, Abmessungen und Gefälle von Entwässerungsrinnen", "es": "fontanería y climatización a arquitectura: los requisitos de posiciones, dimensiones y pendientes de los canales de drenaje", "pt": "hidráulica e climatização para arquitetura: os requisitos de posições, dimensões e declividades das canaletas de drenagem", "ru": "водоснабжение и отопление и вентиляция — архитектуре: требования к расположению, размерам и уклонам водоотводных лотков", "ar": "والسباكة والتدفئة والتهوية إلى العمارة: متطلبات المواقع والأبعاد والميول لقنوات الصرف", "it": "idrici e HVAC a architettura: i requisiti di posizioni, dimensioni e pendenze dei canali di drenaggio", "vi": "cấp thoát nước, điều hòa gửi kiến trúc: yêu cầu vị trí, kích thước, độ dốc của rãnh thoát nước", "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: น้ำหนักบรรทุกบนพื้นจากท่อแขวนและอุปกรณ์ ขนาดฐานราก น้ำหนักอุปกรณ์ และข้อกำหนดที่เกี่ยวข้องของอุปกรณ์ขนาดใหญ่ที่ต้องออกแบบโครงสร้าง ตำแหน่งและแรงดันของขายึดท่อคงที่ในสาขาประปาและระบบปรับอากาศที่ต้องเสริมโครงสร้าง ประปาและระบบปรับอากาศส่งให้สถาปัตยกรรม: ข้อกำหนดตำแหน่ง ขนาด และความลาดเอียงของรางระบายน้ำ ประปาส่งให้สถาปัตยกรรม: ข้อกำหนดของห้องพิเศษ (เช่น ห้องดับเพลิงด้วยแก๊ส ฯลฯ) ต่อช่องเปิด เช่น ประตูและหน้าต่าง"},
    "水专业提建筑：特殊房间（如气体消防房间等）对建筑门、窗等洞口的要求": {"en": "plumbing to architectural: the requirements of special rooms (such as gas fire protection rooms, etc.) for architectural openings such as doors and windows", "ja": "水専門提建築：特殊室（ガス消防室など）の建築ドア・窓などの開口に対する要求", "ko": "급배수가 건축에: 특수실(가스 소방실 등)의 건축 문·창 등 개구부에 대한 요구", "fr": "plomberie à architecture : les exigences des locaux spéciaux (tels que les locaux de protection incendie par gaz, etc.) concernant les baies telles que portes et fenêtres", "de": "Sanitär an Architektur: die Anforderungen besonderer Räume (z. B. Gaslöschanlagenräume) an Öffnungen wie Türen und Fenster", "es": "fontanería a arquitectura: los requisitos de los locales especiales (como los locales de protección contra incendios por gas, etc.) sobre los huecos de puertas y ventanas", "pt": "hidráulica para arquitetura: os requisitos dos ambientes especiais (como salas de proteção contra incêndio por gás, etc.) sobre vãos como portas e janelas", "ru": "водоснабжение — архитектуре: требования особых помещений (например, газового пожаротушения) к проёмам, дверям и окнам", "ar": "والسباكة إلى العمارة: متطلبات الغرف الخاصة (مثل غرف إطفاء الحرائق بالغاز وغيرها) بشأن الفتحات مثل الأبواب والنوافذ", "it": "idrici a architettura: i requisiti dei locali speciali (come le sale di protezione antincendio a gas, ecc.) sulle aperture come porte e finestre", "vi": "cấp thoát nước gửi kiến trúc: yêu cầu của phòng đặc biệt (như phòng chữa cháy khí, v.v.) đối với lỗ mở cửa, cửa sổ của kiến trúc", "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: น้ำหนักบรรทุกบนพื้นจากท่อแขวนและอุปกรณ์ ขนาดฐานราก น้ำหนักอุปกรณ์ และข้อกำหนดที่เกี่ยวข้องของอุปกรณ์ขนาดใหญ่ที่ต้องออกแบบโครงสร้าง ตำแหน่งและแรงดันของขายึดท่อคงที่ในสาขาประปาและระบบปรับอากาศที่ต้องเสริมโครงสร้าง ประปาและระบบปรับอากาศส่งให้สถาปัตยกรรม: ข้อกำหนดตำแหน่ง ขนาด และความลาดเอียงของรางระบายน้ำ ประปาส่งให้สถาปัตยกรรม: ข้อกำหนดของห้องพิเศษ (เช่น ห้องดับเพลิงด้วยแก๊ส ฯลฯ) ต่อช่องเปิด เช่น ประตูและหน้าต่าง"},
    "水提暖：由动力专业提供热源时提出估算总热负荷和二次水温度参数（热媒为蒸汽时提出凝结水回水方式）": {"en": "plumbing to HVAC: when the heat source is provided by the utility discipline, raise the estimated total heat load and the secondary water temperature parameters (when the heating medium is steam, raise the condensate return method)", "ja": "水提暖：動力専門から熱源供給時、推定総熱負荷と二次水温度パラメータを提出する（熱媒が蒸気の場合、復水の還水方式を提出する）", "ko": "급배수가 난방에: 동력 전문이 열원을 공급할 때 추정 총 열부하와 2차수 온도 파라미터를 제출합니다(열매체가 증기일 때 응축수 회수 방식을 제출합니다)", "fr": "plomberie à CVC : lorsque la source de chaleur est fournie par la discipline énergétique, formuler la charge thermique totale estimée et les paramètres de température de l'eau secondaire (si le fluide caloporteur est de la vapeur, formuler le mode de retour des condensats)", "de": "Sanitär an Heizung-Lüftung: bei Wärmequelle durch die Energiedisziplin werden geschätzte Gesamtwärmelast und Sekundärwasser-Temperaturparameter angegeben (bei Dampf als Wärmeträger wird die Kondensatrückführungsart angegeben)", "es": "fontanería a climatización: cuando la fuente de calor la suministre la especialidad de energía, formular la carga térmica total estimada y los parámetros de temperatura del agua secundaria (si el fluido caloportador es vapor, formular el modo de retorno de condensados)", "pt": "hidráulica para climatização: quando a fonte de calor for fornecida pela especialidade de energia, apresentar a carga térmica total estimada e os parâmetros de temperatura da água secundária (se o fluido térmico for vapor, apresentar o modo de retorno de condensado)", "ru": "водоснабжение — отоплению и вентиляции: при подаче тепла от энергетической специальности представить расчётную суммарную тепловую нагрузку и параметры температуры вторичной воды (при паровом теплоносителе — способ возврата конденсата)", "ar": "والسباكة إلى التدفئة والتهوية: عند توفير مصدر الحرارة من تخصص الطاقة، اقتراح الحمل الحراري الإجمالي التقديري ومعايير درجة حرارة الماء الثانوي (إذا كان وسيط التسخين بخاراً، اقتراح طريقة إرجاع المكثفات)", "it": "idrici a HVAC: quando la fonte di calore è fornita dalla disciplina energetica, proporre il carico termico totale stimato e i parametri di temperatura dell'acqua secondaria (se il fluido termovettore è vapore, proporre la modalità di ritorno della condensa)", "vi": "cấp thoát nước gửi điều hòa: khi nguồn nhiệt do chuyên ngành năng lượng cung cấp thì nêu phụ tải nhiệt tổng ước tính và thông số nhiệt độ nước thứ cấp (khi môi chất là hơi nước thì nêu phương thức hồi nước ngưng)", "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น"},
    "暖提电：排烟阀、加压送风阀、需要电气控制的防火阀以及自动控制阀的位置及设备和附件的控制要求": {"en": "HVAC to electrical: the positions of smoke exhaust valves, pressurization air supply valves, fire dampers requiring electrical control and automatic control valves, and the control requirements of equipment and accessories", "ja": "暖提電：排煙弁、加圧送風弁、電気制御が必要な防火弁および自動制御弁の位置並びに設備と付属品の制御要求", "ko": "난방이 전기에: 배연밸브, 가압 송풍밸브, 전기 제어가 필요한 방화밸브 및 자동 제어밸브의 위치 및 설비와 부속품의 제어 요구", "fr": "CVC à électricité : les positions des clapets de désenfumage, des vannes de soufflage en surpression, des clapets coupe-feu à commande électrique et des vannes de régulation automatique, et les exigences de commande des équipements et accessoires", "de": "Heizung-Lüftung an Elektro: die Positionen von Rauchabzugsklappen, Druckbelüftungsklappen, elektrisch gesteuerten Brandschutzklappen und automatischen Regelventilen sowie die Steuerungsanforderungen von Anlagen und Zubehör", "es": "climatización a electricidad: las posiciones de las compuertas de evacuación de humos, compuertas de impulsión de aire a presión, compuertas cortafuego de accionamiento eléctrico y válvulas de control automático, y los requisitos de control de los equipos y accesorios", "pt": "climatização para elétrica: as posições das comportas de exaustão de fumaça, comportas de insuflação de ar pressurizado, comportas corta-fogo de acionamento elétrico e válvulas de controle automático, e os requisitos de controle dos equipamentos e acessórios", "ru": "отопление и вентиляция — электроснабжению: расположение противопожарных клапанов дымоудаления, клапанов подпора воздуха, огнезадерживающих клапанов с электроприводом и автоматических регулирующих клапанов, а также требования к управлению оборудованием и арматурой", "ar": "والتدفئة والتهوية إلى الكهرباء: مواقع صمامات إخراج الدخان وصمامات تزويد الهواء المضغوط وصمامات الحريق التي تتطلب تحكماً كهربائياً وصمامات التحكم الآلي، ومتطلبات التحكم في المعدات والملحقات", "it": "HVAC a elettrici: le posizioni delle serrande di evacuazione fumi, delle serrande di immissione aria in pressione, delle serrande tagliafuoco a comando elettrico e delle valvole di controllo automatico, e i requisiti di controllo di apparecchiature e accessori", "vi": "điều hòa gửi điện: vị trí van xả khói, van cấp gió tăng áp, van chống cháy cần điều khiển điện và van điều khiển tự động, cùng yêu cầu điều khiển thiết bị, phụ kiện", "th": "ประปาส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปาส่งให้ไฟฟ้า: ตำแหน่งหัวจ่ายน้ำดับเพลิง เครื่องแสดงการไหลของน้ำ วาล์วสัญญาณเตือน วาล์วควบคุม ฯลฯ และข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ระบบปรับอากาศส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด และเติมเต็มตำแหน่งและขนาดของช่องระบายอากาศแบบใบพัดภายนอกทั้งหมด (ช่องลมเข้า ช่องลมออก ฯลฯ) ระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งวาล์วระบายควัน วาล์วส่งลมอัดแรงดัน วาล์วกันไฟที่ต้องควบคุมด้วยไฟฟ้า และวาล์วควบคุมอัตโนมัติ รวมถึงข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ไฟฟ้าส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งที่แม่นยำของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนัก พื้น และคาน"},
    "智能化提暖：各弱电机房、安防控制室温湿度参数和通风要求": {"en": "intelligent systems to HVAC: temperature and humidity parameters and ventilation requirements of weak-current equipment rooms and security control rooms", "ja": "智能化→暖：各弱電機械室・安防控制室の温湿度パラメータと換気要求", "ko": "지능화→난방: 각 약전 기계실·보안 통제실의 온습도 파라미터와 환기 요구", "fr": "systèmes intelligents vers CVC : paramètres de température et d'humidité et exigences de ventilation des locaux techniques courants faibles et des salles de contrôle de sécurité", "de": "Gebäudeautomation an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter sowie Lüftungsanforderungen von Schwachstrom-Räumen und Sicherheitskontrollräumen", "es": "sistemas inteligentes a climatización: parámetros de temperatura y humedad y requisitos de ventilación de los cuartos de corrientes débiles y salas de control de seguridad", "pt": "sistemas inteligentes para climatização: parâmetros de temperatura e umidade e requisitos de ventilação das salas de corrente fraca e salas de controle de segurança", "ru": "интеллектуальные системы — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции слаботочных аппаратных и помещений охраны", "ar": "والأنظمة الذكية إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية لغرف التيارات الضعيفة وغرف التحكم الأمنية", "it": "sistemi intelligenti a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali debole corrente e delle sale controllo sicurezza", "vi": "hệ thống thông minh gửi điều hòa: thông số nhiệt độ độ ẩm và yêu cầu thông gió của các phòng máy yếu điện và phòng điều khiển an ninh", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่มีการทำความร้อน ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ปริมาณความร้อนจากอุปกรณ์ในห้องที่มีการปลดปล่อยความร้อนทางไฟฟ้าสูง (หรือกำลังไฟฟ้าติดตั้งโดยประมาณ เช่น ห้องสถานีย่อย ห้องคอมพิวเตอร์เมนเฟรมขนาดใหญ่) ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องอุปกรณ์ไฟฟ้าอ่อนและห้องควบคุมระบบรักษาความปลอดภัย ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณการใช้ไฟฟ้าของห้องอุปกรณ์ไฟฟ้าอ่อน ห้องควบคุมระบบรักษาความปลอดภัย และปล่องไฟฟ้าอ่อน"},
    "电提水：特殊电气用房消防要求": {"en": "electrical to plumbing: fire protection requirements for special electrical rooms", "ja": "電提水：特殊電気室の消防要求を提出する", "ko": "전기가 급배수에: 특수 전기실의 소방 요구를 제출합니다", "fr": "électricité à plomberie : les exigences de protection incendie des locaux électriques spéciaux", "de": "Elektro an Sanitär: Brandschutzanforderungen für besondere Elektroräume", "es": "electricidad a fontanería: requisitos de protección contra incendios de los locales eléctricos especiales", "pt": "elétrica para hidráulica: requisitos de proteção contra incêndio dos ambientes elétricos especiais", "ru": "электроснабжение — водоснабжению: требования пожарной безопасности к специальным электротехническим помещениям", "ar": "والكهرباء إلى السباكة: متطلبات الحماية من الحرائق للغرف الكهربائية الخاصة", "it": "elettrici a idrici: requisiti di protezione antincendio dei locali elettrici speciali", "vi": "điện gửi cấp thoát nước: yêu cầu phòng cháy của phòng điện đặc biệt", "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น"},
    "电提水：特殊房间消防要求": {"en": "electrical to plumbing: fire protection requirements for special rooms", "ja": "電提水：特殊な部屋の消防要求", "ko": "전기가 급배수에: 특수실 소방 요구", "fr": "électricité à plomberie : exigences de protection incendie des pièces spéciales", "de": "Elektro an Sanitär: Brandschutzanforderungen für besondere Räume", "es": "electricidad a fontanería: requisitos de protección contra incendios de los locales especiales", "pt": "elétrica para hidráulica: requisitos de proteção contra incêndio de ambientes especiais", "ru": "электроснабжение — водоснабжению: требования пожарной безопасности к специальным помещениям", "ar": "والكهرباء إلى السباكة: متطلبات الحماية من الحرائق للغرف الخاصة", "it": "elettrici a idrici: requisiti di protezione antincendio dei locali speciali", "vi": "điện gửi cấp thoát nước: yêu cầu phòng cháy phòng đặc biệt", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่ทำความร้อน ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องพิเศษ ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณไฟฟ้าที่ใช้ของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง"},
    "水提建筑：补充完善确认水泵房、水处理机房、热交换站、热水机房、水箱间、报警阀室等机房面积尺寸净高要求，集中管道井位置尺寸，水池尺寸与平面定位、地下室集水坑位置尺寸，影响室内布置和外立面的设备（消火栓、冷却塔等）初步位置，建筑最底层垫层厚度要求": {"en": "Plumbing to architectural: supplement, improve and confirm the area, dimensions and clear height requirements of pump rooms, water treatment rooms, heat exchange stations, hot water rooms, water tank rooms, alarm valve rooms and other plant rooms, the position and dimensions of central riser shafts, the dimensions and plan location of water tanks, the position and dimensions of basement sump pits, the preliminary positions of equipment affecting interior layout and elevations (fire hydrants, cooling towers, etc.), and the thickness requirement of the lowest floor cushion layer", "ja": "水提建築：ポンプ室・水処理室・熱交換ステーション・温水室・水槽室・警報弁室等の機械室面積・寸法・天井高要求を補完・改善・確認し、集中パイプシャフトの位置寸法、水槽の寸法と平面位置、地下集水桝の位置寸法、室内配置と外立面に影響する設備（消火栓・冷却塔等）の初期位置、建築最下層の捨てコンクリート厚さ要求を提出する", "ko": "급배수가 건축에: 펌프실·수처리실·열교환 스테이션·온수실·물탱크실·경보밸브실 등 기계실의 면적·치수·천장고 요구를 보완·개선·확인하고, 중앙 파이프 샤프트 위치 치수, 물탱크 치수와 평면 위치, 지하 집수정 위치 치수, 실내 배치와 외관에 영향을 주는 설비(소화전·냉각탑 등)의 초기 위치, 건축 최하층 바닥 콘크리트 두께 요구를 제출합니다", "fr": "Plomberie à architecture : compléter, améliorer et confirmer les exigences de surface, de dimensions et de hauteur libre des locaux techniques tels que stations de pompage, locaux de traitement d'eau, postes d'échange thermique, locaux d'eau chaude, locaux de réservoirs d'eau, locaux de vannes d'alarme, les positions et dimensions des gaines techniques centrales, les dimensions et la localisation en plan des réservoirs d'eau, les positions et dimensions des puisards de sous-sol, les positions préliminaires des équipements influant sur l'aménagement intérieur et les façades (poteaux d'incendie, tours de refroidissement, etc.) et l'exigence d'épaisseur de la couche de forme du niveau le plus bas", "de": "Sanitär an Architektur: Ergänzung, Verbesserung und Bestätigung der Anforderungen an Fläche, Abmessungen und lichte Höhe von Pumpenräumen, Wasseraufbereitungsräumen, Wärmeübergabestationen, Warmwasserzentralen, Behälterräumen, Alarmventilräumen usw., der Position und Abmessungen zentraler Schächte, der Abmessungen und Grundrisslage von Wasserbehältern, der Position und Abmessungen von Kellerschächten, der vorläufigen Positionen von Anlagen, die Innenanordnung und Fassade beeinflussen (Hydranten, Kühltürme usw.), und der Anforderung an die Dicke der untersten Bodenschicht", "es": "Fontanería a arquitectura: completar, mejorar y confirmar los requisitos de superficie, dimensiones y altura libre de locales técnicos como salas de bombas, salas de tratamiento de agua, estaciones de intercambio térmico, salas de agua caliente, salas de depósitos de agua, salas de válvulas de alarma, etc., las posiciones y dimensiones de los patinillos técnicos centrales, las dimensiones y la ubicación en planta de los depósitos de agua, las posiciones y dimensiones de los sumideros de sótano, las posiciones preliminares de los equipos que afectan a la distribución interior y a las fachadas (hidrantes, torres de refrigeración, etc.) y el requisito de espesor de la capa base del nivel más bajo", "pt": "Hidráulica para arquitetura: complementar, melhorar e confirmar os requisitos de área, dimensões e pé-direito de salas técnicas como salas de bombas, salas de tratamento de água, centrais de troca térmica, salas de água quente, salas de reservatórios, salas de válvulas de alarme, etc., as posições e dimensões dos shaft técnicos centrais, as dimensões e a localização em planta dos reservatórios, as posições e dimensões dos poços de captação do subsolo, as posições preliminares dos equipamentos que afetam a disposição interna e as fachadas (hidrantes, torres de resfriamento, etc.) e a exigência de espessura da camada de base do nível mais baixo", "ru": "Водоснабжение — архитектуре: дополнить, уточнить и подтвердить требования к площадям, размерам и высотам помещений насосных, водоподготовки, теплообменных пунктов, горячего водоснабжения, водяных баков, пожарных клапанных и т.п., положение и размеры центральных шахт, размеры и плановую привязку баков, положение и размеры приямков в подвале, предварительное расположение оборудования, влияющего на внутреннюю планировку и фасады (пожарные гидранты, градирни и т.п.), требование к толщине выравнивающего слоя самого нижнего этажа", "ar": "السباكة إلى العمارة: استكمال وتحسين وتأكيد متطلبات المساحة والأبعاد والارتفاع الصافي للغرف الفنية مثل غرف المضخات وغرف معالجة المياه ومحطات التبادل الحراري وغرف الماء الساخن وغرف خزانات المياه وغرف صمامات الإنذار وغيرها، ومواقع وأبعاد الأعمدة الفنية المركزية، وأبعاد وموقع خزانات المياه في المخطط، ومواقع وأبعاد أحواض التجميع في الطابق السفلي، والمواقع الأولية للمعدات التي تؤثر على التوزيع الداخلي والواجهات (صنابير الإطفاء وأبراج التبريد وغيرها)، ومتطلب سمك طبقة التأسيس لأدنى طابق", "it": "Impianti idrici ad architettura: integrare, migliorare e confermare i requisiti di area, dimensioni e altezza libera di locali tecnici quali sale pompe, sale trattamento acqua, stazioni di scambio termico, sale acqua calda, sale serbatoi, sale valvole di allarme, ecc., le posizioni e dimensioni dei vani tecnici centrali, le dimensioni e la posizione in pianta dei serbatoi, le posizioni e dimensioni dei pozzetti di raccolta in seminterrato, le posizioni preliminari delle apparecchiature che influiscono sulla disposizione interna e sulle facciate (idranti, torri di raffreddamento, ecc.) e il requisito di spessore dello strato di base del livello più basso", "vi": "Cấp thoát nước gửi kiến trúc: bổ sung hoàn thiện và xác nhận yêu cầu diện tích, kích thước, chiều cao thông thủy của phòng bơm, phòng xử lý nước, trạm trao đổi nhiệt, phòng nước nóng, phòng bể nước, phòng van báo động, v.v., vị trí kích thước giếng kỹ thuật tập trung, kích thước và định vị mặt bằng bể nước, vị trí kích thước hố thu nước tầng hầm, vị trí sơ bộ của thiết bị ảnh hưởng bố trí nội thất và mặt ngoài (họng cứu hỏa, tháp giải nhiệt, v.v.), yêu cầu chiều dày lớp lót tầng thấp nhất", "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น"},
    "水、暖、电、智能化提结构：位于楼板上的机房设备和冷却塔的结构荷载；最大长边尺寸或直径≥1000mm的承重结构墙和楼板的预留洞初步位置": {"en": "Plumbing, HVAC, electrical and intelligent systems to structural: the structural loads of machine room equipment and cooling towers located on floor slabs；the preliminary positions of reserved openings with maximum long-side dimension or diameter ≥1000mm in load-bearing structural walls and floor slabs", "ja": "水・暖・電・智能化提構造：床版上の機室設備と冷却塔の構造荷重；最大長辺寸法または直径≥1000mmの耐力構造壁と床版の予備開口初期位置", "ko": "급배수·난방·전기·지능화가 구조에: 바닥 슬래브 위 기계실 설비와 냉각탑의 구조 하중；최대 장변 치수 또는 직경 ≥1000mm의 내력 구조벽과 슬래브 예비 구멍의 초기 위치", "fr": "Plomberie, CVC, électricité et systèmes intelligents à structure : les charges structurelles des équipements de locaux techniques et des tours de refroidissement situés sur les dalles；les positions préliminaires des réservations dont le plus grand côté ou diamètre est ≥ 1000 mm dans les murs porteurs et les dalles", "de": "Sanitär, Heizung-Lüftung, Elektro und Gebäudeautomation an Tragwerk: die Tragwerkslasten der auf Geschossdecken befindlichen Technikraumanlagen und Kühltürme；die vorläufigen Positionen von Durchbrüchen mit maximaler Längsseite oder Durchmesser ≥ 1000 mm in tragenden Wänden und Geschossdecken", "es": "Fontanería, climatización, electricidad y sistemas inteligentes a estructuras: las cargas estructurales de los equipos de los locales técnicos y de las torres de refrigeración situados sobre los forjados；las posiciones preliminares de los huecos cuyo lado mayor o diámetro sea ≥ 1000 mm en muros portantes y forjados", "pt": "Hidráulica, climatização, elétrica e sistemas inteligentes para estruturas: as cargas estruturais dos equipamentos de ambientes técnicos e das torres de resfriamento localizados sobre lajes；as posições preliminares das aberturas com maior lado ou diâmetro ≥ 1000 mm em paredes estruturais e lajes", "ru": "Водоснабжение, отопление и вентиляция, электроснабжение и интеллектуальные системы — конструкциям: нагрузки на конструкции от оборудования технических помещений и градирен, расположенных на перекрытиях；предварительное расположение отверстий с наибольшим размером стороны или диаметром ≥ 1000 мм в несущих стенах и плитах перекрытий", "ar": "السباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية إلى الإنشاءات: الأحمال الإنشائية لمعدات غرف الآلات وأبراج التبريد الموجودة على الألواح；والمواقع الأولية للفتحات التي يبلغ أكبر بعد أو قطر لها ≥ 1000 مم في الجدران الإنشائية الحاملة والألواح", "it": "Idrici, HVAC, elettrici e sistemi intelligenti a strutture: i carichi strutturali delle apparecchiature dei locali tecnici e delle torri di raffreddamento situate sui solai；le posizioni preliminari delle aperture con lato maggiore o diametro ≥ 1000 mm nelle pareti strutturali portanti e nei solai", "vi": "Cấp thoát nước, điều hòa, điện, hệ thống thông minh gửi kết cấu: tải trọng kết cấu của thiết bị phòng kỹ thuật và tháp giải nhiệt đặt trên sàn；vị trí sơ bộ lỗ chờ có cạnh dài lớn nhất hoặc đường kính ≥1000mm trên tường chịu lực và sàn", "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: น้ำหนักบรรทุกโครงสร้างของอุปกรณ์ห้องเครื่องและหอหล่อเย็นที่อยู่บนพื้น ตำแหน่งเบื้องต้นของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนักและพื้น ประปาและระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่ง ระดับแรงดันไฟฟ้า และกำลังไฟฟ้าติดตั้งของอุปกรณ์ไฟฟ้าแรงต่าง ๆ (อุปกรณ์ดับเพลิงและอุปกรณ์ใช้ปกติเสนอแยกกัน)"},
    "水提暖：特殊房间温湿度参数和通风要求（如水箱间温度要求、气体消防房间名称等）": {"en": "Plumbing to HVAC: temperature and humidity parameters and ventilation requirements of special rooms (such as temperature requirements of water tank rooms, names of gas fire-protection rooms, etc.)", "ja": "水→暖：特殊室の温湿度パラメータと換気要求（水槽室の温度要求、ガス系消防室の名称等）", "ko": "급배수→난방: 특수실의 온습도 파라미터와 환기 요구(예: 수조실 온도 요구, 가스 소방실 명칭 등)", "fr": "Plomberie vers CVC : paramètres de température et d'humidité et exigences de ventilation des locaux spéciaux (tels que les exigences de température des locaux de réservoirs d'eau, les noms des locaux de protection incendie à gaz, etc.)", "de": "Sanitär an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter sowie Lüftungsanforderungen besonderer Räume (z. B. Temperaturanforderungen von Wasserbehälterräumen, Bezeichnungen von Gaslöschräumen usw.)", "es": "Fontanería a climatización: parámetros de temperatura y humedad y requisitos de ventilación de locales especiales (como requisitos de temperatura de cuartos de tanques de agua, nombres de locales de protección contra incendios por gas, etc.)", "pt": "Hidráulica para climatização: parâmetros de temperatura e umidade e requisitos de ventilação de ambientes especiais (como requisitos de temperatura de salas de reservatórios de água, nomes de ambientes de proteção contra incêndio por gás, etc.)", "ru": "Водоснабжение — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции специальных помещений (например, требования к температуре помещений водонапорных баков, наименования помещений газового пожаротушения и т.п.)", "ar": "السباكة إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية للغرف الخاصة (مثل متطلبات درجة حرارة غرف خزانات المياه، وأسماء غرف الإطفاء بالغاز وغيرها)", "it": "Impianti idrici a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali speciali (come i requisiti di temperatura dei locali serbatoi acqua, i nomi dei locali di protezione antincendio a gas, ecc.)", "vi": "Cấp thoát nước gửi điều hòa không khí: thông số nhiệt độ độ ẩm và yêu cầu thông gió của phòng đặc biệt (như yêu cầu nhiệt độ phòng bể nước, tên phòng chữa cháy bằng khí, v.v.)", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่มีการทำความร้อน ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ปริมาณความร้อนจากอุปกรณ์ในห้องที่มีการปลดปล่อยความร้อนทางไฟฟ้าสูง (หรือกำลังไฟฟ้าติดตั้งโดยประมาณ เช่น ห้องสถานีย่อย ห้องคอมพิวเตอร์เมนเฟรมขนาดใหญ่) ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องอุปกรณ์ไฟฟ้าอ่อนและห้องควบคุมระบบรักษาความปลอดภัย ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณการใช้ไฟฟ้าของห้องอุปกรณ์ไฟฟ้าอ่อน ห้องควบคุมระบบรักษาความปลอดภัย และปล่องไฟฟ้าอ่อน"},
    "暖提建筑：补充完善确认冷冻机房、水泵房、锅炉房、热交换站、燃气调压站（室）等机房面积尺寸净高，各层分散空调通风机房位置面积，集中管道井位置尺寸，锅炉房等泄爆井（泄爆面）初步位置面积": {"en": "HVAC to architectural: supplement, improve and confirm the area, dimensions and clear height of refrigeration rooms, pump rooms, boiler rooms, heat exchange stations, gas pressure regulating stations (rooms), etc., the positions and areas of decentralized air-conditioning and ventilation rooms on each floor, the position and dimensions of central riser shafts, and the preliminary position and area of relief shafts (relief surfaces) of boiler rooms, etc.", "ja": "暖提建築：冷凍室・ポンプ室・ボイラー室・熱交換ステーション・ガス調圧ステーション（室）等の機械室面積・寸法・天井高を補完・改善・確認し、各階の分散空調換気室の位置面積、集中パイプシャフトの位置寸法、ボイラー室等の爆発逃がしシャフト（逃がし面）の初期位置面積を提出する", "ko": "난방이 건축에: 냉동실·펌프실·보일러실·열교환 스테이션·가스 압력조정소(실) 등 기계실의 면적·치수·천장고를 보완·개선·확인하고, 각 층 분산 공조 환기실의 위치 면적, 중앙 파이프 샤프트 위치 치수, 보일러실 등 폭발 방출 샤프트(방출면)의 초기 위치 면적을 제출합니다", "fr": "CVC à architecture : compléter, améliorer et confirmer la surface, les dimensions et la hauteur libre des locaux de réfrigération, stations de pompage, chaufferies, postes d'échange thermique, postes (locaux) de détente de gaz, etc., les positions et surfaces des locaux décentralisés de CVC et de ventilation à chaque niveau, les positions et dimensions des gaines techniques centrales, la position et la surface préliminaires des conduits (surfaces) de décompression des chaufferies, etc.", "de": "Heizung-Lüftung an Architektur: Ergänzung, Verbesserung und Bestätigung der Fläche, Abmessungen und lichten Höhe von Kälteräumen, Pumpenräumen, Heizräumen, Wärmeübergabestationen, Gasdruckregelstationen (Räumen) usw., der Positionen und Flächen dezentraler Klima- und Lüftungsräume je Geschoss, der Position und Abmessungen zentraler Schächte, der vorläufigen Position und Fläche von Druckentlastungsschächten (Druckentlastungsflächen) der Heizräume usw.", "es": "climatización a arquitectura: completar, mejorar y confirmar la superficie, las dimensiones y la altura libre de salas de refrigeración, salas de bombas, salas de calderas, estaciones de intercambio térmico, estaciones (salas) de regulación de gas, etc., las posiciones y superficies de los locales descentralizados de climatización y ventilación en cada nivel, las posiciones y dimensiones de los patinillos técnicos centrales, la posición y superficie preliminares de los conductos (superficies) de alivio de las salas de calderas, etc.", "pt": "climatização para arquitetura: complementar, melhorar e confirmar a área, as dimensões e o pé-direito de salas de refrigeração, salas de bombas, casas de caldeiras, centrais de troca térmica, estações (salas) de regulagem de gás, etc., as posições e áreas dos ambientes descentralizados de climatização e ventilação em cada pavimento, as posições e dimensões dos shaft técnicos centrais, a posição e a área preliminares dos dutos (superfícies) de alívio das casas de caldeiras, etc.", "ru": "отопление и вентиляция — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты холодильных, насосных, котельных, теплообменных пунктов, газорегуляторных пунктов (помещений) и т.п., расположение и площади децентрализованных вентиляционных камер на каждом этаже, положение и размеры центральных шахт, предварительное расположение и площадь предохранительных шахт (сбросных поверхностей) котельных и т.п.", "ar": "والتدفئة والتهوية إلى العمارة: استكمال وتحسين وتأكيد مساحة وأبعاد وارتفاع غرف التبريد وغرف المضخات وغرف الغلايات ومحطات التبادل الحراري ومحطات (غرف) تنظيم ضغط الغاز وغيرها، ومواقع ومساحات غرف التكييف والتهوية الموزعة في كل طابق، ومواقع وأبعاد الأعمدة الفنية المركزية، والموقع والمساحة الأوليان لأعمدة (أسطح) تفريغ الضغط في غرف الغلايات وغيرها", "it": "HVAC ad architettura: integrare, migliorare e confermare l'area, le dimensioni e l'altezza libera di sale refrigerazione, sale pompe, centrali termiche, stazioni di scambio termico, stazioni (sale) di regolazione del gas, ecc., le posizioni e aree dei locali decentralizzati di climatizzazione e ventilazione a ogni piano, le posizioni e dimensioni dei vani tecnici centrali, la posizione e l'area preliminari dei condotti (superfici) di scarico pressione delle centrali termiche, ecc.", "vi": "điều hòa gửi kiến trúc: bổ sung hoàn thiện xác nhận diện tích, kích thước, chiều cao thông thủy phòng lạnh, phòng bơm, phòng lò hơi, trạm trao đổi nhiệt, trạm (phòng) điều áp gas, v.v., vị trí diện tích phòng điều hòa thông gió phân tán từng tầng, vị trí kích thước giếng kỹ thuật tập trung, vị trí diện tích sơ bộ giếng (bề mặt) thoát áp của phòng lò hơi, v.v.", "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น"},
    "水、暖、电提结构：位于楼板上的机房设备和冷却塔的结构荷载": {"en": "Plumbing, HVAC and electrical to structural: structural loads of equipment rooms on floor slabs and of cooling towers", "ja": "水・暖・電→構造：スラブ上の機械室設備と冷却塔の構造荷重", "ko": "급배수·난방·전기→구조: 슬래브 위 기계실 설비와 냉각탑의 구조 하중", "fr": "Plomberie, CVC et électricité vers structure : charges structurelles des équipements de locaux techniques sur dalles et des tours de refroidissement", "de": "Sanitär, Heizung-Lüftung und Elektro an Tragwerk: Strukturlasten der Geräte in Technikräumen auf Deckenplatten und von Kühltürmen", "es": "Fontanería, climatización y electricidad a estructuras: cargas estructurales de los equipos de locales técnicos sobre losas y de las torres de refrigeración", "pt": "Hidráulica, climatização e elétrica para estruturas: cargas estruturais dos equipamentos de salas técnicas sobre lajes e das torres de resfriamento", "ru": "Водоснабжение, отопление и вентиляция, электроснабжение — конструкциям: конструктивные нагрузки от оборудования технических помещений на плитах перекрытий и от градирен", "ar": "السباكة والتدفئة والتهوية والكهرباء إلى الإنشاءات: الأحمال الإنشائية لمعدات الغرف الفنية على الألواح ولأبراج التبريد", "it": "Impianti idrici, HVAC ed elettrici a strutture: carichi strutturali delle apparecchiature dei locali tecnici su solette e delle torri di raffreddamento", "vi": "Cấp thoát nước, điều hòa, điện gửi kết cấu: tải trọng kết cấu của thiết bị phòng máy trên sàn và tháp giải nhiệt", "th": "ประปา ระบบปรับอากาศ และไฟฟ้า ส่งให้โครงสร้าง: น้ำหนักบรรทุกโครงสร้างของอุปกรณ์ห้องเครื่องบนพื้นแผ่นและหอหล่อเย็น ประปาและระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งของอุปกรณ์ไฟฟ้าแรงทั้งหมดและกำลังไฟฟ้าติดตั้ง (อุปกรณ์ดับเพลิงและอุปกรณ์ใช้งานปกติต้องแจ้งแยกกัน)"},
    "水、暖提电：各强电设备的位置及其电气安装容量（消防设备和平时使用的设备分别提出）": {"en": "plumbing and HVAC to electrical: positions of all high-voltage/strong-current equipment and their electrical installation capacities (fire-protection equipment and normally used equipment shall be stated separately)", "ja": "水・暖→電：各強電設備の位置およびその電気設備容量（消防設備と常用設備は分けて提出する）", "ko": "급배수·난방→전기: 각 강전 설비의 위치 및 전기 설비 용량(소방 설비와 평상시 사용 설비는 각각 제출)", "fr": "plomberie et CVC vers électricité : emplacements des équipements de courant fort et leurs capacités électriques installées (les équipements de protection incendie et les équipements d'usage courant sont indiqués séparément)", "de": "Sanitär und Heizung-Lüftung an Elektro: Positionen aller Starkstromgeräte und ihre elektrischen Anschlussleistungen (Brandschutzgeräte und normal genutzte Geräte sind getrennt anzugeben)", "es": "fontanería y climatización a electricidad: posiciones de todos los equipos de corriente fuerte y sus capacidades eléctricas instaladas (los equipos de protección contra incendios y los de uso habitual se indican por separado)", "pt": "hidráulica e climatização para elétrica: posições de todos os equipamentos de corrente forte e suas capacidades elétricas instaladas (equipamentos de proteção contra incêndio e equipamentos de uso normal devem ser indicados separadamente)", "ru": "водоснабжение и отопление и вентиляция — электроснабжению: расположение всего силового электрооборудования и его установленная электрическая мощность (противопожарное оборудование и оборудование обычного использования указываются раздельно)", "ar": "والسباكة والتدفئة والتهوية إلى الكهرباء: مواقع جميع معدات التيار القوي وسعاتها الكهربائية المركبة (تُذكر معدات الحماية من الحرائق ومعدات الاستخدام العادي بشكل منفصل)", "it": "impianti idrici e HVAC a elettrici: posizioni di tutte le apparecchiature di forza e relative potenze elettriche installate (le apparecchiature antincendio e quelle di uso normale vanno indicate separatamente)", "vi": "cấp thoát nước, điều hòa gửi điện: vị trí các thiết bị điện mạnh và công suất điện lắp đặt của chúng (thiết bị chữa cháy và thiết bị sử dụng thường xuyên nêu riêng)", "th": "ประปา ระบบปรับอากาศ และไฟฟ้า ส่งให้โครงสร้าง: น้ำหนักบรรทุกโครงสร้างของอุปกรณ์ห้องเครื่องบนพื้นแผ่นและหอหล่อเย็น ประปาและระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งของอุปกรณ์ไฟฟ้าแรงทั้งหมดและกำลังไฟฟ้าติดตั้ง (อุปกรณ์ดับเพลิงและอุปกรณ์ใช้งานปกติต้องแจ้งแยกกัน)"},
    "智能化提电：各弱电机房、安防控制室用电量": {"en": "intelligent systems to electrical: power consumption of each weak-current room and security control room", "ja": "智能化提電：各弱電室・安防控制室の電力消費量", "ko": "지능화가 전기에: 각 약전실·안보 통제실의 전력 소비량", "fr": "systèmes intelligents à électricité : consommation électrique de chaque local courants faibles et de la salle de contrôle de sécurité", "de": "Gebäudeautomation an Elektro: Stromverbrauch jedes Schwachstromraums und des Sicherheitskontrollraums", "es": "sistemas inteligentes a electricidad: consumo eléctrico de cada local de corrientes débiles y de la sala de control de seguridad", "pt": "sistemas inteligentes para elétrica: consumo de energia de cada ambiente de corrente fraca e da sala de controle de segurança", "ru": "интеллектуальные системы — электроснабжению: электропотребление слаботочных помещений и диспетчерской охраны", "ar": "والأنظمة الذكية إلى الكهرباء: استهلاك الكهرباء لغرف التيارات الضعيفة وغرفة التحكم الأمنية", "it": "sistemi intelligenti a elettrici: consumo elettrico di ogni locale debole corrente e della sala di controllo sicurezza", "vi": "hệ thống thông minh gửi điện: lượng điện tiêu thụ của các phòng điện yếu, phòng điều khiển an ninh", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่ทำความร้อน ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องพิเศษ ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณไฟฟ้าที่ใช้ของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง"},
    "电提暖：特殊房间温湿度参数和通风要求": {"en": "electrical to HVAC: temperature and humidity parameters and ventilation requirements of special rooms", "ja": "電提暖：特殊な部屋の温湿度パラメータと換気要求", "ko": "전기가 난방에: 특수실 온습도 파라미터와 환기 요구", "fr": "électricité à CVC : paramètres de température et d'humidité et exigences de ventilation des pièces spéciales", "de": "Elektro an Heizung-Lüftung: Temperatur- und Feuchtigkeitsparameter und Lüftungsanforderungen besonderer Räume", "es": "electricidad a climatización: parámetros de temperatura y humedad y requisitos de ventilación de los locales especiales", "pt": "elétrica para climatização: parâmetros de temperatura e umidade e requisitos de ventilação de ambientes especiais", "ru": "электроснабжение — отоплению и вентиляции: параметры температуры и влажности и требования к вентиляции специальных помещений", "ar": "والكهرباء إلى التدفئة والتهوية: معايير درجة الحرارة والرطوبة ومتطلبات التهوية للغرف الخاصة", "it": "elettrici a HVAC: parametri di temperatura e umidità e requisiti di ventilazione dei locali speciali", "vi": "điện gửi điều hòa: thông số nhiệt độ độ ẩm và yêu cầu thông gió phòng đặc biệt", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่ทำความร้อน ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องพิเศษ ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณไฟฟ้าที่ใช้ของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง"},
    "电提建筑：补齐全部管道井尺寸和位置要求": {"en": "electrical to architectural: complete the size and position requirements of all pipe shafts", "ja": "電提建築：すべてのパイプシャフト寸法と位置要求を補完", "ko": "전기가 건축에: 모든 파이프 샤프트 치수와 위치 요구를 보완", "fr": "électricité à architecture : compléter les exigences de dimensions et de positions de toutes les gaines techniques", "de": "Elektro an Architektur: die Größen- und Positionsanforderungen aller Installationsschächte vervollständigen", "es": "electricidad a arquitectura: completar los requisitos de dimensiones y posiciones de todos los patinillos", "pt": "elétrica para arquitetura: completar os requisitos de dimensões e posições de todos os shafts", "ru": "электроснабжение — архитектуре: дополнить требования к размерам и расположению всех инженерных шахт", "ar": "والكهرباء إلى العمارة: استكمال متطلبات الأبعاد والمواقع لجميع الأعمدة الفنية", "it": "elettrici a architettura: completare i requisiti di dimensioni e posizioni di tutti i cavedi", "vi": "điện gửi kiến trúc: bổ sung đầy đủ yêu cầu kích thước và vị trí các hộp kỹ thuật đường ống", "th": "ประปาส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปาส่งให้ไฟฟ้า: ตำแหน่งหัวจ่ายน้ำดับเพลิง เครื่องแสดงการไหลของน้ำ วาล์วสัญญาณเตือน วาล์วควบคุม ฯลฯ และข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ระบบปรับอากาศส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด และเติมเต็มตำแหน่งและขนาดของช่องระบายอากาศแบบใบพัดภายนอกทั้งหมด (ช่องลมเข้า ช่องลมออก ฯลฯ) ระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งวาล์วระบายควัน วาล์วส่งลมอัดแรงดัน วาล์วกันไฟที่ต้องควบคุมด้วยไฟฟ้า และวาล์วควบคุมอัตโนมัติ รวมถึงข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ไฟฟ้าส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งที่แม่นยำของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนัก พื้น และคาน"},
    "暖提电：需要防雷接地和防静电的设备名称（油气锅炉房、烟囱、燃气放散管、可燃气体管道等）": {"en": "HVAC to electrical: names of equipment requiring lightning protection, earthing and anti-static measures (oil and gas boiler rooms, chimneys, gas vent pipes, combustible gas pipelines, etc.)", "ja": "暖提電：防雷接地・防静電を要する設備名（油ガスボイラー室・煙突・ガス放散管・可燃性ガス配管等）", "ko": "난방이 전기에: 방뢰 접지·정전기 방지가 필요한 설비명(유가스 보일러실·굴뚝·가스 방산관·가연성 가스 배관 등)", "fr": "CVC à électricité : noms des équipements nécessitant une mise à la terre de protection contre la foudre et antistatique (chaufferies gazole/gaz, cheminées, conduits de vente de gaz, canalisations de gaz combustible, etc.)", "de": "Heizung-Lüftung an Elektro: Namen der Anlagen mit Anforderungen an Blitzschutzerdung und Antistatik (Öl- und Gasheizräume, Schornsteine, Gasabblaseleitungen, brennbare Gasleitungen usw.)", "es": "climatización a electricidad: nombres de los equipos que requieren puesta a tierra de protección contra rayos y protección antiestática (salas de calderas de gasóleo/gas, chimeneas, tubos de venteo de gas, tuberías de gas combustible, etc.)", "pt": "climatização para elétrica: nomes dos equipamentos que exigem aterramento contra raios e proteção antiestática (casas de caldeiras a óleo/gás, chaminés, tubos de vente de gás, tubulações de gás combustível, etc.)", "ru": "отопление и вентиляция — электроснабжению: наименования оборудования, требующего молниезащитного заземления и антистатической защиты (мазутно-газовые котельные, дымовые трубы, газовые свечи рассеивания, трубопроводы горючих газов и т.п.)", "ar": "والتدفئة والتهوية إلى الكهرباء: أسماء المعدات التي تتطلب تأريض الحماية من الصواعق والحماية من الكهرباء الساكنة (غرف الغلايات النفطية/الغازية والمداخن وأنابيب تصريف الغاز وخطوط الغاز القابلة للاشتعال وغيرها)", "it": "HVAC a elettrici: nomi delle apparecchiature che richiedono messa a terra contro i fulmini e protezione antistatica (centrali a olio/gas, camini, tubi di sfogo gas, tubazioni di gas combustibile, ecc.)", "vi": "điều hòa gửi điện: tên thiết bị cần tiếp đất chống sét và chống tĩnh điện (phòng lò hơi dầu khí, ống khói, ống xả khí gas, đường ống khí cháy được, v.v.)", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่ทำความร้อน ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องพิเศษ ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณไฟฟ้าที่ใช้ของห้องไฟฟ้าอ่อนและห้องควบคุมความปลอดภัยแต่ละห้อง"},
    "暖提水：主要用水点房间名称和估算用水量、主要排水点房间名称和排水性质、冷却水总流量进出水温要求和拟定冷水机组台数": {"en": "HVAC to plumbing: the room names and estimated water consumption of main water use points, the room names and drainage nature of main drainage points, the total cooling water flow rate, inlet/outlet water temperature requirements and the proposed number of chiller units", "ja": "暖提水：主要用水点の室名と推定用水量、主要排水点の室名と排水性質、冷却水総流量・出入口水温要求と予定冷凍機台数を提出する", "ko": "난방이 급배수에: 주요 용수점 실명과 추정 용수량, 주요 배수점 실명과 배수 성질, 냉각수 총 유량·입출구 수온 요구와 예정 냉동기 대수를 제출합니다", "fr": "CVC à plomberie : les noms des pièces des principaux points d'utilisation d'eau et l'estimation de la consommation d'eau, les noms des pièces des principaux points de drainage et la nature du drainage, le débit total de l'eau de refroidissement, les exigences de température d'entrée/sortie et le nombre proposé de groupes refroidisseurs", "de": "Heizung-Lüftung an Sanitär: Raumnamen und geschätzter Wasserverbrauch der Hauptwasserentnahmepunkte, Raumnamen und Entwässerungsart der Hauptentwässerungspunkte, Gesamtdurchfluss des Kühlwassers, Vor- und Rücklauftemperaturanforderungen und die geplante Anzahl der Kältemaschinen", "es": "climatización a fontanería: los nombres de las salas de los puntos principales de consumo de agua y el consumo estimado, los nombres de las salas de los puntos principales de drenaje y la naturaleza del drenaje, el caudal total del agua de refrigeración, los requisitos de temperatura de entrada/salida y el número propuesto de grupos enfriadores", "pt": "climatização para hidráulica: os nomes dos ambientes dos principais pontos de consumo de água e o consumo estimado, os nomes dos ambientes dos principais pontos de drenagem e a natureza da drenagem, a vazão total da água de resfriamento, os requisitos de temperatura de entrada/saída e o número proposto de chillers", "ru": "отопление и вентиляция — водоснабжению: наименования помещений основных точек водоразбора и расчётные расходы воды, наименования помещений основных точек водоотведения и характер водоотведения, общий расход охлаждающей воды, требования к температурам на входе и выходе и планируемое количество холодильных машин", "ar": "والتدفئة والتهوية إلى السباكة: أسماء غرف نقاط استهلاك الماء الرئيسية واستهلاك الماء التقديري، وأسماء غرف نقاط الصرف الرئيسية وطبيعة الصرف، ومعدل التدفق الإجمالي لمياه التبريد ومتطلبات درجة حرارة الدخول/الخروج وعدد وحدات التبريد المقترح", "it": "HVAC a idrici: i nomi dei locali dei principali punti di prelievo acqua e il consumo stimato, i nomi dei locali dei principali punti di scarico e la natura dello scarico, la portata totale dell'acqua di raffreddamento, i requisiti di temperatura ingresso/uscita e il numero proposto di gruppi frigoriferi", "vi": "điều hòa gửi cấp thoát nước: tên phòng điểm dùng nước chính và lượng nước dùng ước tính, tên phòng điểm thoát nước chính và tính chất thoát nước, tổng lưu lượng nước làm mát, yêu cầu nhiệt độ nước vào/ra và số máy lạnh dự kiến", "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น"},
    "电、智能化提总图：市政管线（电力、通讯、有线电视等）至本楼接口大致位置，主要室外管线布置及构筑物（检修井等）大致位置": {"en": "electrical and intelligent systems to site planning: approximate positions of the interfaces of municipal pipelines (power, communications, cable TV, etc.) to this building, approximate positions of major outdoor pipeline layouts and structures (inspection wells, etc.)", "ja": "電・智能化提総図：市政管線（電力、通信、有線テレビなど）から本棟への接続おおよその位置、主要屋外配管配置及び構築物（点検井など）のおおよその位置", "ko": "전기·지능화가 총도에: 시정 관선(전력, 통신, 유선TV 등)의 본 건물 접속 대략 위치, 주요 옥외 배관 배치 및 구조물(점검 우물 등) 대략 위치", "fr": "électricité et systèmes intelligents à l'aménagement du site : positions approximatives des interfaces des réseaux municipaux (électricité, télécommunications, télévision par câble, etc.) avec ce bâtiment, positions approximatives des principaux réseaux extérieurs et ouvrages (regards de visite, etc.)", "de": "Elektro und Gebäudeautomation an Standortplanung: ungefähre Positionen der Anschlusspunkte der städtischen Leitungen (Strom, Kommunikation, Kabelfernsehen usw.) an dieses Gebäude, ungefähre Positionen der Hauptaußenleitungen und Bauwerke (Inspektionsschächte usw.)", "es": "electricidad y sistemas inteligentes a urbanización: posiciones aproximadas de los puntos de conexión de las redes municipales (electricidad, comunicaciones, televisión por cable, etc.) con este edificio, posiciones aproximadas de las principales redes exteriores y obras (arquetas de registro, etc.)", "pt": "elétrica e sistemas inteligentes para urbanismo: posições aproximadas dos pontos de conexão das redes municipais (energia, comunicações, TV a cabo, etc.) com este edifício, posições aproximadas dos principais traçados externos e obras (poços de inspeção, etc.)", "ru": "электроснабжение и интеллектуальные системы — генплану: приблизительное расположение точек присоединения городских линий (электроснабжение, связь, кабельное телевидение и т.п.) к данному зданию, приблизительное расположение основных наружных линий и сооружений (смотровые колодцы и т.п.)", "ar": "والكهرباء والأنظمة الذكية إلى تخطيط المواقع: المواقع التقريبية لنقاط ربط خطوط المرافق البلدية (الكهرباء، الاتصالات، التلفزيون الكبلي وغيرها) بهذا المبنى، والمواقع التقريبية للخطوط الخارجية الرئيسية والمنشآت (آبار التفتيش وغيرها)", "it": "elettrici e sistemi intelligenti a pianificazione del sito: posizioni approssimative dei punti di collegamento delle linee comunali (energia, comunicazioni, TV via cavo, ecc.) con questo edificio, posizioni approssimative dei principali tracciati esterni e delle opere (pozzetti di ispezione, ecc.)", "vi": "điện, hệ thống thông minh gửi tổng mặt bằng: vị trí tương đối điểm đấu nối đường ống hạ tầng (điện, viễn thông, truyền hình cáp, v.v.) vào công trình, vị trí tương đối bố trí đường ống ngoài chính và công trình (giếng kiểm tra, v.v.)", "th": "ประปาส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อเครือข่ายสาธารณูปโภคต่าง ๆ (ประปา ท่อระบายน้ำ น้ำร้อน ฯลฯ) เข้าสู่อาคารนี้ ตำแหน่งโดยประมาณของการวางท่อภายนอกหลักและโครงสร้างหลัก (ถังบำบัด ถังดักไขมัน บ่อมาตรวัดน้ำ ข้อต่อปั๊มดับเพลิง ฯลฯ) ระบบปรับอากาศส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อเครือข่ายสาธารณูปโภค (ความร้อน แก๊ส) เข้าสู่อาคาร ตำแหน่งโดยประมาณของการวางท่อภายนอกหลักและโครงสร้าง (บ่อวาล์ว ฯลฯ) ไฟฟ้าและระบบอัจฉริยะส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อสายสาธารณูปโภค (ไฟฟ้า การสื่อสาร เคเบิลทีวี ฯลฯ) เข้าสู่อาคาร ตำแหน่งโดยประมาณของการวางสายภายนอกหลักและโครงสร้าง (บ่อตรวจสอบ ฯลฯ)"},
    "水、暖、电、智能化提结构：最大长边尺寸或直径≥1000mm的承重墙、楼板及梁上预留洞的准确位置": {"en": "plumbing, HVAC, electrical and intelligent systems to structural: the accurate positions of reserved openings with maximum long-side dimension or diameter ≥1000mm in load-bearing walls, floor slabs and beams", "ja": "水・暖・電・智能化提構造：最大長辺寸法または直径≥1000mmの耐力壁、床版及び梁上の予備開口の正確な位置", "ko": "급배수·난방·전기·지능화가 구조에: 최대 장변 치수 또는 직경 ≥1000mm의 내력벽, 슬래브 및 보 위 예비 구멍의 정확한 위치", "fr": "plomberie, CVC, électricité et systèmes intelligents à structure : les positions précises des réservations dont le plus grand côté ou diamètre est ≥ 1000 mm dans les murs porteurs, les dalles et les poutres", "de": "Sanitär, Heizung-Lüftung, Elektro und Gebäudeautomation an Tragwerk: die genauen Positionen von Durchbrüchen mit maximaler Längsseite oder Durchmesser ≥ 1000 mm in tragenden Wänden, Geschossdecken und Balken", "es": "fontanería, climatización, electricidad y sistemas inteligentes a estructuras: las posiciones precisas de los huecos cuyo lado mayor o diámetro sea ≥ 1000 mm en muros portantes, forjados y vigas", "pt": "hidráulica, climatização, elétrica e sistemas inteligentes para estruturas: as posições precisas das aberturas com maior lado ou diâmetro ≥ 1000 mm em paredes estruturais, lajes e vigas", "ru": "водоснабжение, отопление и вентиляция, электроснабжение и интеллектуальные системы — конструкциям: точное расположение отверстий с наибольшим размером стороны или диаметром ≥ 1000 мм в несущих стенах, плитах перекрытий и балках", "ar": "والسباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية إلى الإنشاءات: المواقع الدقيقة للفتحات التي يبلغ أكبر بعد أو قطر لها ≥ 1000 مم في الجدران الحاملة والألواح والكمرات", "it": "idrici, HVAC, elettrici e sistemi intelligenti a strutture: le posizioni precise delle aperture con lato maggiore o diametro ≥ 1000 mm nelle pareti portanti, nei solai e nelle travi", "vi": "cấp thoát nước, điều hòa, điện, hệ thống thông minh gửi kết cấu: vị trí chính xác các lỗ chờ có cạnh dài lớn nhất hoặc đường kính ≥1000mm trên tường chịu lực, sàn và dầm", "th": "ประปาส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปาส่งให้ไฟฟ้า: ตำแหน่งหัวจ่ายน้ำดับเพลิง เครื่องแสดงการไหลของน้ำ วาล์วสัญญาณเตือน วาล์วควบคุม ฯลฯ และข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ระบบปรับอากาศส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด และเติมเต็มตำแหน่งและขนาดของช่องระบายอากาศแบบใบพัดภายนอกทั้งหมด (ช่องลมเข้า ช่องลมออก ฯลฯ) ระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งวาล์วระบายควัน วาล์วส่งลมอัดแรงดัน วาล์วกันไฟที่ต้องควบคุมด้วยไฟฟ้า และวาล์วควบคุมอัตโนมัติ รวมถึงข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ไฟฟ้าส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งที่แม่นยำของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนัก พื้น และคาน"},
    "暖提水：不采暖房间名称": {"en": "HVAC to plumbing: names of unheated rooms", "ja": "暖→水：非暖房室の名称", "ko": "난방→급배수: 비난방실 명칭", "fr": "CVC vers plomberie : noms des locaux non chauffés", "de": "Heizung-Lüftung an Sanitär: Bezeichnungen unbeheizter Räume", "es": "climatización a fontanería: nombres de locales sin calefacción", "pt": "climatização para hidráulica: nomes de ambientes sem aquecimento", "ru": "отопление и вентиляция — водоснабжению: наименования неотапливаемых помещений", "ar": "والتدفئة والتهوية إلى السباكة: أسماء الغرف غير المدفأة", "it": "HVAC a impianti idrici: nomi dei locali non riscaldati", "vi": "điều hòa gửi cấp thoát nước: tên phòng không sưởi", "th": "ประปาส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ (เช่น ข้อกำหนดอุณหภูมิห้องถังน้ำ ชื่อห้องดับเพลิงด้วยแก๊ส ฯลฯ) ระบบปรับอากาศส่งให้ประปา: ชื่อห้องที่ไม่มีการทำความร้อน ไฟฟ้าส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องพิเศษ ปริมาณความร้อนจากอุปกรณ์ในห้องที่มีการปลดปล่อยความร้อนทางไฟฟ้าสูง (หรือกำลังไฟฟ้าติดตั้งโดยประมาณ เช่น ห้องสถานีย่อย ห้องคอมพิวเตอร์เมนเฟรมขนาดใหญ่) ระบบอัจฉริยะส่งให้ระบบปรับอากาศ: ค่าอุณหภูมิความชื้นและข้อกำหนดการระบายอากาศของห้องอุปกรณ์ไฟฟ้าอ่อนและห้องควบคุมระบบรักษาความปลอดภัย ระบบอัจฉริยะส่งให้ไฟฟ้า: ปริมาณการใช้ไฟฟ้าของห้องอุปกรณ์ไฟฟ้าอ่อน ห้องควบคุมระบบรักษาความปลอดภัย และปล่องไฟฟ้าอ่อน"},
    "水提电：消火栓、水流指示器、报警阀、控制阀等的位置及设备和附件的控制要求": {"en": "plumbing to electrical: the positions of fire hydrants, water flow indicators, alarm valves, control valves, etc., and the control requirements of equipment and accessories", "ja": "水提電：消火栓、水流指示器、警報弁、制御弁などの位置並びに設備と付属品の制御要求", "ko": "급배수가 전기에: 소화전, 수류 지시기, 경보밸브, 제어밸브 등의 위치 및 설비와 부속품의 제어 요구", "fr": "plomberie à électricité : les positions des poteaux d'incendie, indicateurs de débit, vannes d'alarme, vannes de régulation, etc., et les exigences de commande des équipements et accessoires", "de": "Sanitär an Elektro: die Positionen von Löschhydranten, Strömungswächtern, Alarmventilen, Regelventilen usw. und die Steuerungsanforderungen von Anlagen und Zubehör", "es": "fontanería a electricidad: las posiciones de las bocas de incendio, indicadores de flujo, válvulas de alarma, válvulas de control, etc., y los requisitos de control de los equipos y accesorios", "pt": "hidráulica para elétrica: as posições dos hidrantes, indicadores de fluxo, válvulas de alarme, válvulas de controle, etc., e os requisitos de controle dos equipamentos e acessórios", "ru": "водоснабжение — электроснабжению: расположение пожарных кранов, сигнализаторов потока, сигнальных клапанов, регулирующих клапанов и т.п., а также требования к управлению оборудованием и арматурой", "ar": "والسباكة إلى الكهرباء: مواقع صنابير الإطفاء ومؤشرات التدفق وصمامات الإنذار وصمامات التحكم وغيرها، ومتطلبات التحكم في المعدات والملحقات", "it": "idrici a elettrici: le posizioni di idranti, indicatori di flusso, valvole di allarme, valvole di controllo, ecc., e i requisiti di controllo di apparecchiature e accessori", "vi": "cấp thoát nước gửi điện: vị trí họng cứu hỏa, chỉ báo dòng chảy, van báo động, van điều khiển, v.v. và yêu cầu điều khiển thiết bị, phụ kiện", "th": "ประปาส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปาส่งให้ไฟฟ้า: ตำแหน่งหัวจ่ายน้ำดับเพลิง เครื่องแสดงการไหลของน้ำ วาล์วสัญญาณเตือน วาล์วควบคุม ฯลฯ และข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ระบบปรับอากาศส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด และเติมเต็มตำแหน่งและขนาดของช่องระบายอากาศแบบใบพัดภายนอกทั้งหมด (ช่องลมเข้า ช่องลมออก ฯลฯ) ระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งวาล์วระบายควัน วาล์วส่งลมอัดแรงดัน วาล์วกันไฟที่ต้องควบคุมด้วยไฟฟ้า และวาล์วควบคุมอัตโนมัติ รวมถึงข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ไฟฟ้าส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งที่แม่นยำของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนัก พื้น และคาน"},
    "暖提总图：市政管网（热力、燃气）至本楼接口大致位置，主要室外管道布置及构筑物（阀门井等）大致位置": {"en": "HVAC to site planning: approximate positions of the interfaces of municipal pipeline networks (heating, gas) to this building, approximate positions of major outdoor pipeline layouts and structures (valve wells, etc.)", "ja": "暖提総図：市政管網（熱力、ガス）から本棟への接続おおよその位置、主要屋外配管配置及び構築物（バルブ井など）のおおよその位置", "ko": "난방이 총도에: 시정 관망(열력, 가스)의 본 건물 접속 대략 위치, 주요 옥외 배관 배치 및 구조물(밸브 우물 등) 대략 위치", "fr": "CVC à l'aménagement du site : positions approximatives des interfaces des réseaux municipaux (chauffage, gaz) avec ce bâtiment, positions approximatives des principaux réseaux extérieurs et ouvrages (regards de vannes, etc.)", "de": "Heizung-Lüftung an Standortplanung: ungefähre Positionen der Anschlusspunkte der städtischen Netze (Fernwärme, Gas) an dieses Gebäude, ungefähre Positionen der Hauptaußenleitungen und Bauwerke (Absperrschieberschächte usw.)", "es": "climatización a urbanización: posiciones aproximadas de los puntos de conexión de las redes municipales (calefacción, gas) con este edificio, posiciones aproximadas de las principales redes exteriores y obras (arquetas de válvulas, etc.)", "pt": "climatização para urbanismo: posições aproximadas dos pontos de conexão das redes municipais (aquecimento, gás) com este edifício, posições aproximadas dos principais traçados externos e obras (caixas de válvulas, etc.)", "ru": "отопление и вентиляция — генплану: приблизительное расположение точек присоединения городских сетей (теплоснабжение, газ) к данному зданию, приблизительное расположение основных наружных сетей и сооружений (задвижки в колодцах и т.п.)", "ar": "والتدفئة والتهوية إلى تخطيط المواقع: المواقع التقريبية لنقاط ربط شبكات المرافق البلدية (الحرارة، الغاز) بهذا المبنى، والمواقع التقريبية للشبكات الخارجية الرئيسية والمنشآت (آبار الصمامات وغيرها)", "it": "HVAC a pianificazione del sito: posizioni approssimative dei punti di collegamento delle reti comunali (teleriscaldamento, gas) con questo edificio, posizioni approssimative dei principali tracciati esterni e delle opere (pozzetti valvole, ecc.)", "vi": "điều hòa gửi tổng mặt bằng: vị trí tương đối điểm đấu nối mạng hạ tầng (nhiệt, gas) vào công trình, vị trí tương đối bố trí đường ống ngoài nhà chính và công trình (giếng van, v.v.)", "th": "ประปาส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อเครือข่ายสาธารณูปโภคต่าง ๆ (ประปา ท่อระบายน้ำ น้ำร้อน ฯลฯ) เข้าสู่อาคารนี้ ตำแหน่งโดยประมาณของการวางท่อภายนอกหลักและโครงสร้างหลัก (ถังบำบัด ถังดักไขมัน บ่อมาตรวัดน้ำ ข้อต่อปั๊มดับเพลิง ฯลฯ) ระบบปรับอากาศส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อเครือข่ายสาธารณูปโภค (ความร้อน แก๊ส) เข้าสู่อาคาร ตำแหน่งโดยประมาณของการวางท่อภายนอกหลักและโครงสร้าง (บ่อวาล์ว ฯลฯ) ไฟฟ้าและระบบอัจฉริยะส่งให้การวางผังบริเวณ: ตำแหน่งโดยประมาณของจุดเชื่อมต่อสายสาธารณูปโภค (ไฟฟ้า การสื่อสาร เคเบิลทีวี ฯลฯ) เข้าสู่อาคาร ตำแหน่งโดยประมาณของการวางสายภายนอกหลักและโครงสร้าง (บ่อตรวจสอบ ฯลฯ)"},
    "暖提电：需防雷接地和防静电的设备名称（油气锅炉房、烟囱、燃气放散管、可燃气体管道等）": {"en": "HVAC to electrical: the names of equipment requiring lightning protection, earthing and anti-static measures (oil and gas boiler rooms, chimneys, gas vent pipes, combustible gas pipelines, etc.)", "ja": "暖提電：防雷接地・防静電を要する設備名（油ガスボイラー室・煙突・ガス放散管・可燃性ガス配管等）を提出する", "ko": "난방이 전기에: 방뢰 접지·정전기 방지가 필요한 설비명(유가스 보일러실·굴뚝·가스 방산관·가연성 가스 배관 등)을 제출합니다", "fr": "CVC à électricité : les noms des équipements nécessitant une mise à la terre de protection contre la foudre et antistatique (chaufferies gazole/gaz, cheminées, cheminées de ventilation de gaz, canalisations de gaz combustible, etc.)", "de": "Heizung-Lüftung an Elektro: Namen der Anlagen mit Anforderungen an Blitzschutzerdung und Antistatik (Öl- und Gasheizräume, Schornsteine, Gasabblaseleitungen, brennbare Gasleitungen usw.)", "es": "climatización a electricidad: los nombres de los equipos que requieren puesta a tierra de protección contra rayos y protección antiestática (salas de calderas de gasóleo/gas, chimeneas, tubos de venteo de gas, tuberías de gas combustible, etc.)", "pt": "climatização para elétrica: os nomes dos equipamentos que exigem aterramento contra raios e proteção antiestática (casas de caldeiras a óleo/gás, chaminés, tubos de vente de gás, tubulações de gás combustível, etc.)", "ru": "отопление и вентиляция — электроснабжению: наименования оборудования, требующего молниезащитного заземления и антистатической защиты (мазутно-газовые котельные, дымовые трубы, газовые свечи рассеивания, трубопроводы горючих газов и т.п.)", "ar": "والتدفئة والتهوية إلى الكهرباء: أسماء المعدات التي تتطلب تأريض الحماية من الصواعق والحماية من الكهرباء الساكنة (غرف الغلايات النفطية/الغازية والمداخن وأنابيب تصريف الغاز وخطوط الغاز القابلة للاشتعال وغيرها)", "it": "HVAC a elettrici: i nomi delle apparecchiature che richiedono messa a terra contro i fulmini e protezione antistatica (centrali a olio/gas, camini, tubi di sfogo gas, tubazioni di gas combustibile, ecc.)", "vi": "điều hòa gửi điện: tên thiết bị cần tiếp đất chống sét và chống tĩnh điện (phòng lò hơi dầu khí, ống khói, ống xả khí gas, đường ống khí cháy được, v.v.)", "th": "ประปาส่งให้สถาปัตยกรรม: เสริม เติมเต็ม และยืนยันข้อกำหนดพื้นที่ ขนาด ความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊ม ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องน้ำร้อน ห้องถังน้ำ ห้องวาล์วสัญญาณเตือน ฯลฯ ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ขนาดและตำแหน่งในผังของถังน้ำ ตำแหน่งขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่กระทบการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) ข้อกำหนดความหนาชั้นรองพื้นของชั้นล่างสุด ประปาส่งให้ระบบปรับอากาศ: เมื่อแหล่งความร้อนมาจากสาขาพลังงาน ให้เสนอภาระความร้อนรวมโดยประมาณและค่าอุณหภูมิน้ำทุติยภูมิ (ถ้าสื่อความร้อนเป็นไอน้ำ ให้เสนอวิธีส่งคืนน้ำควบแน่น) ระบบปรับอากาศส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องเย็น ห้องปั๊ม ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน สถานี (ห้อง) ควบคุมแรงดันแก๊ส ฯลฯ ตำแหน่งพื้นที่ห้องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งขนาดปล่องเทคนิคส่วนกลาง ตำแหน่งพื้นที่เบื้องต้นของปล่อง (ผิว) ระบายแรงดันของห้องหม้อไอน้ำ ฯลฯ ระบบปรับอากาศส่งให้ประปา: ชื่อห้องจุดใช้น้ำหลักและปริมาณการใช้น้ำโดยประมาณ ชื่อห้องจุดระบายน้ำหลักและลักษณะการระบายน้ำ อัตราการไหลรวมน้ำหล่อเย็น ข้อกำหนดอุณหภูมิน้ำเข้า/ออก และจำนวนเครื่องทำน้ำเย็นที่กำหนด ระบบปรับอากาศส่งให้ไฟฟ้า: ชื่ออุปกรณ์ที่ต้องต่อลงดินป้องกันฟ้าผ่าและป้องกันไฟฟ้าสถิต (ห้องหม้อไอน้ำน้ำมัน/แก๊ส ปล่องไฟ ท่อระบายแก๊ส ท่อแก๊สติดไฟได้ ฯลฯ) ไฟฟ้าส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของสถานีไฟฟ้าย่อย ห้องเครื่องกำเนิดไฟฟ้า ฯลฯ ตำแหน่งพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งขนาดปล่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาพื้น ตำแหน่งพื้นที่ปล่องดูด/ระบายอากาศหล่อเย็นเครื่องกำเนิดไฟฟ้า ไฟฟ้าส่งให้ประปา: ข้อกำหนดอัคคีภัยของห้องไฟฟ้าพิเศษ ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เสริม เติมเต็ม ยืนยันพื้นที่ ขนาด ความสูงโปร่งของห้องสายเข้า ห้องเครือข่าย ห้องชุมสายโทรศัพท์ ห้องเคเบิลทีวี ห้องครอบคลุมสัญญาณมือถือ ห้องควบคุมความปลอดภัย ห้องทีวีดาวเทียม ฯลฯ ตำแหน่งขนาดปล่องไฟฟ้าอ่อน ห้องควบคุมห้องประชุมใหญ่ ข้อกำหนดความหนาพื้น"},
    "水提建筑：补齐全部管道井尺寸和位置要求": {"en": "Plumbing to architectural: complete the size and position requirements of all pipe shafts", "ja": "水提建築：すべてのパイプシャフト寸法と位置要求を補完", "ko": "급배수가 건축에: 모든 파이프 샤프트 치수와 위치 요구를 보완", "fr": "Plomberie à architecture : compléter les exigences de dimensions et de positions de toutes les gaines techniques", "de": "Sanitär an Architektur: die Größen- und Positionsanforderungen aller Installationsschächte vervollständigen", "es": "Fontanería a arquitectura: completar los requisitos de dimensiones y posiciones de todos los patinillos", "pt": "Hidráulica para arquitetura: completar os requisitos de dimensões e posições de todos os shafts", "ru": "Водоснабжение — архитектуре: дополнить требования к размерам и расположению всех инженерных шахт", "ar": "السباكة إلى العمارة: استكمال متطلبات الأبعاد والمواقع لجميع الأعمدة الفنية", "it": "Idrici a architettura: completare i requisiti di dimensioni e posizioni di tutti i cavedi", "vi": "Cấp thoát nước gửi kiến trúc: bổ sung đầy đủ yêu cầu kích thước và vị trí các hộp kỹ thuật đường ống", "th": "ประปาส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปาส่งให้ไฟฟ้า: ตำแหน่งหัวจ่ายน้ำดับเพลิง เครื่องแสดงการไหลของน้ำ วาล์วสัญญาณเตือน วาล์วควบคุม ฯลฯ และข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ระบบปรับอากาศส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด และเติมเต็มตำแหน่งและขนาดของช่องระบายอากาศแบบใบพัดภายนอกทั้งหมด (ช่องลมเข้า ช่องลมออก ฯลฯ) ระบบปรับอากาศส่งให้ไฟฟ้า: ตำแหน่งวาล์วระบายควัน วาล์วส่งลมอัดแรงดัน วาล์วกันไฟที่ต้องควบคุมด้วยไฟฟ้า และวาล์วควบคุมอัตโนมัติ รวมถึงข้อกำหนดการควบคุมอุปกรณ์และส่วนประกอบ ไฟฟ้าส่งให้สถาปัตยกรรม: เติมเต็มข้อกำหนดขนาดและตำแหน่งของช่องเดินท่อทั้งหมด ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งที่แม่นยำของช่องเผื่อที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม. ในผนังรับน้ำหนัก พื้น และคาน"},
    "水、暖、电、智能化提结构：最大长边尺寸或直径≥1000mm的承重结构墙和楼板的预留洞初步位置": {"en": "Plumbing, HVAC, electrical and intelligent systems to structural: the preliminary positions of reserved openings in load-bearing structural walls and floor slabs with a maximum long-side dimension or diameter of ≥1000mm;", "ja": "水・暖・電・智能化から構造専門へ：最大長辺寸法または直径が1000mm以上の耐力構造壁と床スラブの予備開口の初期位置；", "ko": "급배수·난방·전기·지능화가 구조에: 최대 장변 치수 또는 직경이 1000mm 이상인 내력 구조벽과 슬래브의 예비 구멍 초기 위치;", "fr": "Plomberie, CVC, électricité et systèmes intelligents à structure : les positions préliminaires des réservations dans les murs porteurs et les dalles dont la plus grande dimension ou le diamètre est ≥ 1000 mm ;", "de": "Sanitär, Heizung-Lüftung, Elektro und Gebäudeautomation an Tragwerk: die vorläufigen Positionen von Durchbrüchen in tragenden Wänden und Geschossdecken mit maximaler Längsseite oder Durchmesser ≥ 1000 mm;", "es": "Fontanería, climatización, electricidad y sistemas inteligentes a estructuras: las posiciones preliminares de los huecos reservados en muros portantes y forjados cuya mayor dimensión o diámetro sea ≥ 1000 mm;", "pt": "Hidráulica, climatização, elétrica e sistemas inteligentes para estruturas: as posições preliminares das aberturas reservadas em paredes estruturais e lajes cuja maior dimensão ou diâmetro seja ≥ 1000 mm;", "ru": "Водоснабжение, отопление и вентиляция, электроснабжение и интеллектуальные системы — конструктивной специальности: предварительное расположение отверстий в несущих стенах и плитах перекрытий с наибольшим размером стороны или диаметром ≥ 1000 мм;", "ar": "السباكة والتدفئة والتهوية والكهرباء والأنظمة الذكية إلى الإنشاءات: المواقع الأولية للفتحات المخصصة في الجدران الإنشائية الحاملة والألواح التي يبلغ أكبر بعد أو قطرها ≥ 1000 مم؛", "it": "Idrici, HVAC, elettrici e sistemi intelligenti a strutture: le posizioni preliminari delle aperture riservate nelle pareti strutturali portanti e nelle solette con dimensione massima del lato o diametro ≥ 1000 mm;", "vi": "Cấp thoát nước, điều hòa, điện, hệ thống thông minh gửi kết cấu: vị trí sơ bộ các lỗ chờ trên tường chịu lực và sàn có cạnh dài lớn nhất hoặc đường kính ≥1000mm;", "th": "ประปา ระบบปรับอากาศ ไฟฟ้า และระบบอัจฉริยะส่งให้โครงสร้าง: ตำแหน่งเบื้องต้นของช่องเผื่อในผนังโครงสร้างรับน้ำหนักและพื้นที่มีด้านยาวสูงสุดหรือเส้นผ่านศูนย์กลาง ≥ 1000 มม."},
    "水、暖、电提总图：各市政管线（给水、排水、热力、燃气、电力、电信、有线等）在本楼的接口位置要求，主要室外管线布置及主要相关构筑物（化粪池、隔油池、水表井、水泵结合器、阀门井、管线检修井、烟囱、室外箱式变压器等）初步位置，有燃油系统时提出室外储油罐占用面积、初步位置和相关技术要求": {"en": "plumbing, HVAC and electrical to site planning: the interface position requirements of various municipal pipelines (water supply, drainage, heating, gas, power, telecom, cable TV, etc.) entering this building, the preliminary positions of main outdoor pipeline layouts and main related structures (septic tanks, grease interceptors, water meter wells, fire pump connections, valve wells, pipeline inspection wells, chimneys, outdoor pad-mounted transformers, etc.); when there is an oil fuel system, raise the occupied area, preliminary position and related technical requirements of the outdoor fuel storage tank", "ja": "水・暖・電から総図専門へ：各市政配管（給水・排水・熱力・ガス・電力・通信・有線等）の本棟への接続位置要求、主要屋外配管の配置および主要関連構築物（化糞池・グリーストラップ・水道メーター桝・ポンプ接合器・バルブ桝・管路点検桝・煙突・屋外キュービクル変圧器等）の初期位置；燃油システムがある場合、屋外燃料貯蔵タンクの占有面積・初期位置・関連技術要求を提出する", "ko": "급배수·난방·전기가 총도에: 각 시정 배관(급수·배수·열력·가스·전력·통신·유선 등)의 본 건물 접속 위치 요구, 주요 옥외 배관 배치 및 주요 관련 구조물(정화조·기름분리조·수도계량기 우물·소방펌프접합기·밸브 우물·관로 점검 우물·굴뚝·옥외 박스변압기 등)의 초기 위치; 유류 시스템이 있을 때 옥외 유류 저장탱크 점유 면적·초기 위치·관련 기술 요구를 제출합니다", "fr": "plomberie, CVC et électricité à aménagement du site : les exigences de position des raccordements des divers réseaux municipaux (eau potable, assainissement, chauffage urbain, gaz, électricité, télécommunications, câble, etc.) au bâtiment, les positions préliminaires des principaux tracés de réseaux extérieurs et des principales ouvrages associés (fosses septiques, séparateurs de graisses, regards de compteurs d'eau, raccords de pompes incendie, regards de vannes, regards de visite, cheminées, postes de transformation extérieurs, etc.) ; en cas de système au fioul, formuler la surface occupée, la position préliminaire et les exigences techniques du réservoir de stockage de fioul extérieur", "de": "Sanitär, Heizung-Lüftung und Elektro an Standortplanung: die Anforderungen an die Anschlusspositionen der verschiedenen städtischen Leitungen (Wasserversorgung, Entwässerung, Fernwärme, Gas, Strom, Telekom, Kabel usw.) am Gebäude, die vorläufigen Positionen der wichtigsten Außenleitungsführungen und der wichtigsten zugehörigen Bauwerke (Klärgruben, Fettabscheider, Wasserzähler-Schächte, Feuerlöschpumpenanschlüsse, Schieber-Schächte, Leitungsrevisionsschächte, Schornsteine, Ortsnetzstationen im Freien usw.); bei einem Ölfeuerungssystem: belegte Fläche, vorläufige Position und technische Anforderungen des Außen-Öllagerbehälters angeben", "es": "fontanería, climatización y electricidad a urbanización: los requisitos de posición de las conexiones de los distintos redes municipales (abastecimiento, saneamiento, calefacción, gas, electricidad, telecomunicaciones, cable, etc.) al edificio, las posiciones preliminares de los principales trazados de redes exteriores y de las principales obras asociadas (fosas sépticas, separadores de grasas, arquetas de contadores de agua, conexiones de bombas de incendio, arquetas de válvulas, arquetas de registro, chimeneas, centros de transformación exteriores, etc.); en caso de sistema de combustible líquido, formular la superficie ocupada, la posición preliminar y los requisitos técnicos del depósito exterior de almacenamiento de combustible", "pt": "hidráulica, climatização e elétrica para urbanismo: os requisitos de posição das conexões das diversas redes municipais (abastecimento, drenagem, aquecimento, gás, eletricidade, telecomunicações, cabo, etc.) ao edifício, as posições preliminares dos principais traçados de redes externas e das principais obras associadas (fossas sépticas, separadores de gordura, poços de hidrômetros, conexões de bombas de incêndio, poços de válvulas, poços de inspeção, chaminés, subestações externas, etc.); no caso de sistema de óleo combustível, apresentar a área ocupada, a posição preliminar e os requisitos técnicos do tanque externo de armazenamento de combustível", "ru": "водоснабжение, отопление и вентиляция, электроснабжение — специальности генплана: требования к местам присоединения городских сетей (водопровод, канализация, теплоснабжение, газ, электроэнергия, связь, кабельное ТВ и т.п.) к зданию, предварительное расположение основных наружных сетей и основных связанных сооружений (септики, жироуловители, водомерные колодцы, соединительные головки пожарных насосов, задвижечные колодцы, смотровые колодцы, дымовые трубы, наружные блочные трансформаторные подстанции и т.п.); при наличии системы на жидком топливе представить занимаемую площадь, предварительное расположение и технические требования к наружному резервуару хранения топлива", "ar": "والسباكة والتدفئة والتهوية والكهرباء إلى تخطيط الموقع: متطلبات مواقع توصيل الشبكات البلدية المختلفة (إمداد المياه والصرف والتدفئة والغاز والكهرباء والاتصالات والكابل وغيرها) بالمبنى، والمواقع الأولية لتخطيطات الشبكات الخارجية الرئيسية والمنشآت المرتبطة الرئيسية (خزانات الصرف الصحي وفواصل الدهون وآبار عدادات المياه ووصلات مضخات الإطفاء وآبار الصمامات وآبار فحص الخطوط والمداخن ومحولات الصناديق الخارجية وغيرها)؛ وعند وجود نظام وقود سائل، اقتراح المساحة المشغولة والموقع الأولي والمتطلبات الفنية لخزان تخزين الوقود الخارجي", "it": "idrici, HVAC ed elettrici a pianificazione del sito: i requisiti di posizione dei collegamenti delle varie reti comunali (acquedotto, fognatura, teleriscaldamento, gas, energia elettrica, telecomunicazioni, cavo, ecc.) all'edificio, le posizioni preliminari dei principali tracciati di reti esterne e delle principali opere correlate (fosse settiche, separatori di grassi, pozzetti contatori acqua, attacchi pompe antincendio, pozzetti valvole, pozzetti di ispezione, camini, cabine di trasformazione esterne, ecc.); in caso di sistema a olio combustibile, proporre l'area occupata, la posizione preliminare e i requisiti tecnici del serbatoio esterno di stoccaggio del combustibile", "vi": "cấp thoát nước, điều hòa, điện gửi tổng mặt bằng: yêu cầu vị trí đấu nối các tuyến hạ tầng kỹ thuật (cấp nước, thoát nước, nhiệt, gas, điện, viễn thông, truyền hình cáp, v.v.) vào tòa nhà, vị trí sơ bộ bố trí đường ống ngoài nhà chính và các công trình liên quan chính (bể tự hoại, bể tách dầu mỡ, giếng đồng hồ nước, đầu nối bơm cứu hỏa, giếng van, giếng kiểm tra đường ống, ống khói, trạm biến áp ngoài trời, v.v.); khi có hệ thống nhiên liệu dầu thì nêu diện tích chiếm chỗ, vị trí sơ bộ và yêu cầu kỹ thuật liên quan của bể chứa dầu ngoài trời", "th": "ประปา ระบบปรับอากาศ และไฟฟ้าส่งให้การวางผังบริเวณ: ข้อกำหนดตำแหน่งเชื่อมต่อท่อสาธารณูปโภคต่างๆ (น้ำประปา ระบายน้ำ ความร้อน แก๊ส ไฟฟ้า โทรคมนาคม เคเบิล ฯลฯ) เข้าอาคาร ตำแหน่งเบื้องต้นของการวางท่อภายนอกหลักและสิ่งก่อสร้างที่เกี่ยวข้องหลัก (บ่อเกรอะ ถังแยกไขมัน บ่อมิเตอร์น้ำ จุดต่อปั๊มดับเพลิง บ่อวาล์ว บ่อตรวจสอบท่อ ปล่องไฟ หม้อแปลงชนิดตู้นอกอาคาร ฯลฯ) เมื่อมีระบบน้ำมันเชื้อเพลิง ให้เสนอพื้นที่ใช้สอย ตำแหน่งเบื้องต้น และข้อกำหนดทางเทคนิคที่เกี่ยวข้องของถังเก็บน้ำมันภายนอก"},
    "水提建筑：补充完善确认水泵房、水处理机房、热交换站、热水机房、水箱间、报警阀室等机房面积尺寸净高，集中管道井位置尺寸，水池尺寸与平面定位、地下室集水坑位置尺寸，影响建筑室内布置和外立面的设备（消火栓、冷却塔等）初步位置，最底层垫层厚度要求": {"en": "Plumbing to architectural: supplement, improve and confirm the area, dimensions and clear height of machine rooms such as pump rooms, water treatment rooms, heat exchange stations, hot water machine rooms, water tank rooms and alarm valve rooms, the positions and dimensions of centralized pipe shafts, the dimensions and plan positions of pools, the positions and dimensions of basement sumps, the preliminary positions of equipment affecting the interior layout and exterior elevation of the building (fire hydrants, cooling towers, etc.), and the thickness requirements of the lowermost floor screed;", "ja": "水提建築：ポンプ室、水処理室、熱交換ステーション、給湯機室、水槽間、警報弁室などの機械室面積・寸法・天井高、集中パイプシャフトの位置・寸法、水池の寸法と平面位置、地下室集水桝の位置・寸法、建築室内レイアウトと外観立面に影響する設備（消火栓、冷却塔など）の初期位置、最下層敷き均し厚さ要求を補充・改善・確認する；", "ko": "급배수가 건축에: 펌프실, 수처리실, 열교환 스테이션, 급탕기실, 물탱크실, 경보밸브실 등의 기계실 면적·치수·순높이, 집중 파이프 샤프트 위치·치수, 물탱크 치수와 평면 위치, 지하 집수정 위치·치수, 건축 실내 배치와 외관 입면에 영향을 주는 설비(소화전, 냉각탑 등)의 초기 위치, 최하층 되메움(기초 슬래브) 두께 요구를 보완·개선·확인합니다;", "fr": "Plomberie à architecture : compléter, améliorer et confirmer les surfaces, dimensions et hauteurs nettes des locaux techniques tels que les locaux de pompes, les locaux de traitement de l'eau, les postes de livraison de chaleur, les locaux de production d'eau chaude, les locaux de réservoirs et les locaux de vannes d'alarme, les positions et dimensions des gaines techniques centralisées, les dimensions et implantations en plan des bassins, les positions et dimensions des puisards de sous-sol, les positions préliminaires des équipements affectant l'agencement intérieur et l'élévation extérieure du bâtiment (poteaux d'incendie, tours de refroidissement, etc.) et les exigences d'épaisseur de la chape du niveau le plus bas ;", "de": "Sanitär an Architektur: die Flächen, Abmessungen und lichten Höhen von Technikräumen wie Pumpenräumen, Wasseraufbereitungsräumen, Wärmeübergabestationen, Warmwasseranlagenräumen, Behälterräumen und Alarmventilräumen, die Positionen und Abmessungen zentraler Installationsschächte, die Abmessungen und Lagepläne von Becken, die Positionen und Abmessungen von Untergeschoss-Sümpfen, die vorläufigen Positionen von Anlagen, die die Innenanordnung und die Außenansicht des Gebäudes beeinflussen (Löschhydranten, Kühltürme usw.), sowie die Dickeanforderungen des untersten Estrichs ergänzen, verbessern und bestätigen;", "es": "Fontanería a arquitectura: complementar, mejorar y confirmar las superficies, dimensiones y alturas libres de los locales técnicos como salas de bombas, salas de tratamiento de agua, estaciones de intercambio de calor, salas de producción de agua caliente, salas de depósitos de agua y salas de válvulas de alarma, las posiciones y dimensiones de los patinillos técnicos centralizados, las dimensiones y emplazamientos en planta de los depósitos, las posiciones y dimensiones de los sumideros del sótano, las posiciones preliminares de los equipos que afectan a la distribución interior y a la fachada del edificio (bocas de incendio, torres de refrigeración, etc.) y los requisitos de espesor de la solera del nivel inferior;", "pt": "Hidráulica para arquitetura: complementar, aperfeiçoar e confirmar as áreas, dimensões e pés-direitos dos ambientes técnicos como salas de bombas, salas de tratamento de água, centrais de troca de calor, salas de produção de água quente, salas de reservatórios e salas de válvulas de alarme, as posições e dimensões dos shafts técnicos centralizados, as dimensões e posicionamentos em planta dos reservatórios, as posições e dimensões dos poços de coleta do subsolo, as posições preliminares dos equipamentos que afetam a disposição interna e a fachada do edifício (hidrantes, torres de resfriamento, etc.) e os requisitos de espessura da camada de regularização do nível mais baixo;", "ru": "Водоснабжение — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений (насосных, водоподготовки, тепловых пунктов, горячего водоснабжения, водомерных, узлов с сигнальными клапанами и т.п.), положение и размеры центральных инженерных шахт, размеры и привязку в плане резервуаров, положение и размеры приямков в подвале, предварительное расположение оборудования, влияющего на внутреннюю планировку и наружный вид здания (пожарные краны, градирни и т.п.), требования к толщине стяжки самого нижнего уровня;", "ar": "السباكة إلى العمارة: استكمال وتحسين وتأكيد مساحات وأبعاد وارتفاعات غرف الآلات مثل غرف المضخات وغرف معالجة المياه ومحطات التبادل الحراري وغرف الماء الساخن وغرف الخزانات وغرف صمامات الإنذار، ومواقع وأبعاد الأعمدة الفنية المركزية، وأبعاد ومواقع الأحواض في المخطط، ومواقع وأبعاد آبار التجميع في الطابق السفلي، والمواقع الأولية للمعدات التي تؤثر على التوزيع الداخلي والواجهة الخارجية للمبنى (صنابير الإطفاء، أبراج التبريد وغيرها)، ومتطلبات سماكة طبقة التسوية في أدنى مستوى؛", "it": "Idrici a architettura: integrare, migliorare e confermare le aree, le dimensioni e le altezze utili dei locali tecnici come sale pompe, sale trattamento acqua, stazioni di scambio termico, sale produzione acqua calda, sale serbatoi e sale valvole di allarme, le posizioni e dimensioni dei cavedi tecnici centralizzati, le dimensioni e le posizioni planimetriche delle vasche, le posizioni e dimensioni dei pozzetti del seminterrato, le posizioni preliminari delle apparecchiature che influenzano la disposizione interna e l'elevazione esterna dell'edificio (idranti, torri di raffreddamento, ecc.) e i requisiti di spessore del massetto del livello più basso;", "vi": "Cấp thoát nước gửi kiến trúc: bổ sung, hoàn thiện, xác nhận diện tích, kích thước, chiều cao thông thủy các phòng kỹ thuật như phòng bơm, phòng xử lý nước, trạm trao đổi nhiệt, phòng máy nước nóng, phòng bể nước, phòng van báo động; vị trí, kích thước hộp kỹ thuật tập trung; kích thước và định vị mặt bằng bể nước; vị trí, kích thước hố thu nước tầng hầm; vị trí sơ bộ thiết bị ảnh hưởng bố trí nội thất và mặt đứng công trình (họng cứu hỏa, tháp giải nhiệt, v.v.); yêu cầu chiều dày lớp lót tầng dưới cùng;", "th": "ประปาส่งให้สถาปัตยกรรม: เพิ่มเติม ปรับปรุง และยืนยันพื้นที่ ขนาด และความสูงโปร่งของห้องเครื่อง เช่น ห้องปั๊มน้ำ ห้องบำบัดน้ำ สถานีแลกเปลี่ยนความร้อน ห้องเครื่องน้ำร้อน ห้องถังน้ำ และห้องวาล์วสัญญาณเตือน ตำแหน่งและขนาดช่องเดินท่อรวม ขนาดและตำแหน่งผังของบ่อน้ำ ตำแหน่งและขนาดบ่อรวบรวมน้ำใต้ดิน ตำแหน่งเบื้องต้นของอุปกรณ์ที่ส่งผลต่อการจัดวางภายในและภายนอกอาคาร (หัวจ่ายน้ำดับเพลิง หอหล่อเย็น ฯลฯ) และข้อกำหนดความหนาของชั้นปรับระดับชั้นล่างสุด"},
    "暖提建筑：补充完善确认冷冻机房、水泵房、锅炉房、热交换站、燃气调压站（室）等机房面积尺寸净高，各层分散空调通风机房位置面积，集中管道井位置尺寸，锅炉房泄爆井（泄爆面）初步位置面积，地板采暖区域或房间名称及对建筑地面做法厚度要求，室外立面百页窗（新风口、排风口等）初步位置尺寸": {"en": "HVAC to architectural: supplement, improve and confirm the area, dimensions and clear height of machine rooms such as chiller rooms, pump rooms, boiler rooms, heat exchange stations and gas pressure regulating stations (rooms), the positions and areas of decentralized air-conditioning and ventilation machine rooms on each floor, the positions and dimensions of centralized pipe shafts, the preliminary positions and areas of boiler room explosion vent shafts (vent faces), the names of floor heating areas or rooms and the thickness requirements for building floor construction, and the preliminary positions and dimensions of exterior louver windows (fresh air outlets, exhaust outlets, etc.);", "ja": "暖提建築：冷凍機室、ポンプ室、ボイラー室、熱交換ステーション、ガス調圧ステーション（室）などの機械室面積・寸法・天井高、各階の分散空調・換気機室の位置・面積、集中パイプシャフトの位置・寸法、ボイラー室の爆発逃がし井（逃がし面）の初期位置・面積、床暖房区域または部屋名と建築床工法・厚さ要求、室外立面ルーバー窓（新風口、排風口など）の初期位置・寸法を補充・改善・確認する；", "ko": "난방이 건축에: 냉동기실, 펌프실, 보일러실, 열교환 스테이션, 가스 압력조정 스테이션(실) 등의 기계실 면적·치수·순높이, 각 층 분산 공조·환기 기계실 위치·면적, 집중 파이프 샤프트 위치·치수, 보일러실 폭발 배출정(배출면) 초기 위치·면적, 바닥난방 구역 또는 실명과 건축 바닥 공법·두께 요구, 외부 입면 루버 창(신풍구, 배풍구 등) 초기 위치·치수를 보완·개선·확인합니다;", "fr": "CVC à architecture : compléter, améliorer et confirmer les surfaces, dimensions et hauteurs nettes des locaux techniques tels que les locaux des groupes froids, les locaux de pompes, les chaufferies, les postes de livraison de chaleur et les postes (locaux) de détente de gaz, les positions et surfaces des locaux techniques de climatisation et de ventilation décentralisés de chaque niveau, les positions et dimensions des gaines techniques centralisées, les positions et surfaces préliminaires des puits d'évents de soufflage des chaufferies, les noms des zones ou pièces de plancher chauffant et les exigences d'épaisseur des dallages, ainsi que les positions et dimensions préliminaires des fenêtres à persiennes en façade (bouches d'air neuf, bouches d'extraction, etc.) ;", "de": "Heizung-Lüftung an Architektur: die Flächen, Abmessungen und lichten Höhen von Technikräumen wie Kältemaschinenräumen, Pumpenräumen, Heizräumen, Wärmeübergabestationen und Gasdruckregelstationen (-räumen), die Positionen und Flächen dezentraler Klima- und Lüftungstechnikräume je Geschoss, die Positionen und Abmessungen zentraler Installationsschächte, die vorläufigen Positionen und Flächen der Explosionsentlastungsschächte von Heizräumen, die Namen von Fußbodenheizungsbereichen oder -räumen und die Anforderungen an die Fußbodenkonstruktion und -dicke des Gebäudes sowie die vorläufigen Positionen und Abmessungen von außenliegenden Lamellenfenstern (Frischluftöffnungen, Abluftöffnungen usw.);", "es": "climatización a arquitectura: complementar, mejorar y confirmar las superficies, dimensiones y alturas libres de los locales técnicos como salas de enfriadoras, salas de bombas, salas de calderas, estaciones de intercambio de calor y estaciones (locales) de regulación de gas, las posiciones y superficies de los locales técnicos descentralizados de climatización y ventilación de cada planta, las posiciones y dimensiones de los patinillos técnicos centralizados, las posiciones y superficies preliminares de los conductos de alivio de explosión de las salas de calderas, los nombres de las zonas o locales de calefacción por suelo radiante y los requisitos de espesor de los solados, así como las posiciones y dimensiones preliminares de las ventanas de lamas de fachada (tomas de aire nuevo, salidas de extracción, etc.);", "pt": "climatização para arquitetura: complementar, aperfeiçoar e confirmar as áreas, dimensões e pés-direitos dos ambientes técnicos como salas de chillers, salas de bombas, casas de caldeiras, centrais de troca de calor e estações (ambientes) de regulação de gás, as posições e áreas dos ambientes técnicos descentralizados de climatização e ventilação de cada pavimento, as posições e dimensões dos shafts técnicos centralizados, as posições e áreas preliminares dos poços de alívio de explosão das casas de caldeiras, os nomes das áreas ou ambientes de piso radiante e os requisitos de espessura dos revestimentos de piso, bem como as posições e dimensões preliminares das janelas de lâminas na fachada (entradas de ar novo, saídas de exaustão, etc.);", "ru": "отопление и вентиляция — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений (холодильных, насосных, котельных, тепловых пунктов, газорегуляторных пунктов (помещений) и т.п.), расположение и площади децентрализованных венткамер и камер кондиционирования по этажам, положение и размеры центральных инженерных шахт, предварительное расположение и площади взрывных клапанов (разгрузочных поверхностей) котельных, наименования зон или помещений тёплых полов и требования к конструкции и толщине полов, предварительное расположение и размеры фасадных жалюзийных решёток (воздухозаборов, вытяжных отверстий и т.п.);", "ar": "والتدفئة والتهوية إلى العمارة: استكمال وتحسين وتأكيد مساحات وأبعاد وارتفاعات غرف الآلات مثل غرف المبردات وغرف المضخات وغرف الغلايات ومحطات التبادل الحراري ومحطات (غرف) تنظيم الغاز، ومواقع ومساحات غرف التكييف والتهوية الموزعة في كل طابق، ومواقع وأبعاد الأعمدة الفنية المركزية، والمواقع والمساحات الأولية لأعمدة تفريغ الانفجار في غرف الغلايات، وأسماء مناطق أو غرف التدفئة الأرضية ومتطلبات سماكة تنفيذ الأرضيات، والمواقع والأبعاد الأولية لنوافذ الشرائح الخارجية (مداخل الهواء الجديد، مخارج الهواء العادم وغيرها)؛", "it": "HVAC a architettura: integrare, migliorare e confermare le aree, le dimensioni e le altezze utili dei locali tecnici come sale gruppi frigoriferi, sale pompe, centrali termiche, stazioni di scambio termico e stazioni (locali) di regolazione del gas, le posizioni e aree dei locali tecnici decentralizzati di climatizzazione e ventilazione di ogni piano, le posizioni e dimensioni dei cavedi tecnici centralizzati, le posizioni e aree preliminari dei camini di sfogo esplosione delle centrali termiche, i nomi delle zone o dei locali a pavimento radiante e i requisiti di spessore delle pavimentazioni, nonché le posizioni e dimensioni preliminari delle finestre a lamelle in facciata (prese d'aria nuova, bocchette di estrazione, ecc.);", "vi": "điều hòa gửi kiến trúc: bổ sung, hoàn thiện, xác nhận diện tích, kích thước, chiều cao thông thủy các phòng kỹ thuật như phòng máy lạnh, phòng bơm, phòng lò hơi, trạm trao đổi nhiệt, trạm (phòng) điều áp gas; vị trí, diện tích các phòng kỹ thuật điều hòa thông gió phân tán từng tầng; vị trí, kích thước hộp kỹ thuật tập trung; vị trí, diện tích sơ bộ giếng xả nổ phòng lò hơi; tên khu vực hoặc phòng sưởi sàn và yêu cầu chiều dày cấu tạo sàn kiến trúc; vị trí, kích thước sơ bộ cửa chớp mặt đứng (miệng gió tươi, miệng gió thải, v.v.);", "th": "ระบบปรับอากาศส่งให้สถาปัตยกรรม: เพิ่มเติม ปรับปรุง และยืนยันพื้นที่ ขนาด และความสูงโปร่งของห้องเครื่อง เช่น ห้องเครื่องทำความเย็น ห้องปั๊มน้ำ ห้องหม้อไอน้ำ สถานีแลกเปลี่ยนความร้อน และสถานี (ห้อง) ควบคุมแรงดันแก๊ส ตำแหน่งและพื้นที่ห้องเครื่องปรับอากาศระบายอากาศกระจายแต่ละชั้น ตำแหน่งและขนาดช่องเดินท่อรวม ตำแหน่งและพื้นที่เบื้องต้นของช่องระบายแรงดันระเบิดห้องหม้อไอน้ำ ชื่อพื้นที่หรือห้องทำความร้อนพื้นและข้อกำหนดความหนาของพื้นอาคาร และตำแหน่งและขนาดเบื้องต้นของช่องระบายอากาศแบบใบพัดภายนอก (ช่องลมเข้า ช่องลมออก ฯลฯ)"},
    "电提建筑：补充完善确认变配电室、发电机房等机房面积尺寸净高（含变配电室夹层要求），各层弱电机房配电小室位置面积，主要强电竖井位置尺寸，楼板厚度要求，发电机冷却进风与排风竖井位置面积": {"en": "electrical to architectural: supplement, improve and confirm the area, dimensions and clear height of machine rooms such as transformer and distribution rooms and generator rooms (including the mezzanine requirements of transformer and distribution rooms), the positions and areas of weak-current machine rooms and distribution cubicles on each floor, the positions and dimensions of main strong-current shafts, floor thickness requirements, and the positions and areas of generator cooling air inlet and exhaust shafts;", "ja": "電提建築：変配電室、発電機室などの機械室面積・寸法・天井高（変配電室中二階要求を含む）、各階の弱電機室・配電小室の位置・面積、主要強電シャフトの位置・寸法、床版厚さ要求、発電機冷却給気・排気シャフトの位置・面積を補充・改善・確認する；", "ko": "전기가 건축에: 변배전실, 발전기실 등의 기계실 면적·치수·순높이(변배전실 중간층 요구 포함), 각 층 약전 기계실·배전 소실 위치·면적, 주요 강전 샤프트 위치·치수, 바닥 슬래브 두께 요구, 발전기 냉각 급기·배기 샤프트 위치·면적을 보완·개선·확인합니다;", "fr": "électricité à architecture : compléter, améliorer et confirmer les surfaces, dimensions et hauteurs nettes des locaux techniques tels que les postes de transformation et de distribution et les locaux de groupes électrogènes (y compris les exigences de mezzanine des postes de transformation), les positions et surfaces des locaux courants faibles et des petits locaux de distribution de chaque niveau, les positions et dimensions des principales gaines courants forts, les exigences d'épaisseur des planchers et les positions et surfaces des gaines d'air de refroidissement des groupes électrogènes ;", "de": "Elektro an Architektur: die Flächen, Abmessungen und lichten Höhen von Technikräumen wie Trafo- und Verteilungsräumen und Generatorräumen (einschließlich der Anforderungen an Zwischengeschosse der Traforäume), die Positionen und Flächen von Schwachstrom-Technikräumen und Verteilerkabinen je Geschoss, die Positionen und Abmessungen der Hauptstarkstromschächte, die Anforderungen an die Geschossdeckenstärke sowie die Positionen und Flächen der Kühlluftzuluft- und Abluftschächte der Generatoren;", "es": "electricidad a arquitectura: complementar, mejorar y confirmar las superficies, dimensiones y alturas libres de los locales técnicos como centros de transformación y salas de grupos electrógenos (incluidos los requisitos de entreplanta de los centros de transformación), las posiciones y superficies de los locales de corrientes débiles y de los cuartos de distribución de cada planta, las posiciones y dimensiones de los patinillos principales de corrientes fuertes, los requisitos de espesor de forjados y las posiciones y superficies de los conductos de entrada y salida de aire de refrigeración de los grupos electrógenos;", "pt": "elétrica para arquitetura: complementar, aperfeiçoar e confirmar as áreas, dimensões e pés-direitos dos ambientes técnicos como subestações e salas de geradores (incluindo os requisitos de mezanino das subestações), as posições e áreas dos ambientes de corrente fraca e pequenas salas de distribuição de cada pavimento, as posições e dimensões dos shafts principais de corrente forte, os requisitos de espessura de lajes e as posições e áreas dos dutos de entrada e saída de ar de refrigeração dos geradores;", "ru": "электроснабжение — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений (трансформаторных, дизельных и т.п., включая требования к антресолям трансформаторных), расположение и площади слаботочных и распределительных помещений по этажам, положение и размеры основных сильноточных шахт, требования к толщине перекрытий, положение и площади шахт забора и выброса охлаждающего воздуха дизель-генераторов;", "ar": "والكهرباء إلى العمارة: استكمال وتحسين وتأكيد مساحات وأبعاد وارتفاعات غرف الآلات مثل غرف المحولات والتوزيع وغرف المولدات (بما في ذلك متطلبات الطابق النصفي لغرف المحولات)، ومواقع ومساحات غرف التيار الضعيف وحجرات التوزيع في كل طابق، ومواقع وأبعاد أعمدة التيار القوي الرئيسية، ومتطلبات سماكة الألواح، ومواقع ومساحات أعمدة دخول وخروج هواء تبريد المولدات؛", "it": "elettrici a architettura: integrare, migliorare e confermare le aree, le dimensioni e le altezze utili dei locali tecnici come cabine di trasformazione e sale gruppi elettrogeni (inclusi i requisiti di soppalco delle cabine), le posizioni e aree dei locali debole corrente e dei piccoli locali di distribuzione di ogni piano, le posizioni e dimensioni dei cavedi principali di corrente forte, i requisiti di spessore dei solai e le posizioni e aree dei condotti di aspirazione e scarico dell'aria di raffreddamento dei gruppi elettrogeni;", "vi": "điện gửi kiến trúc: bổ sung, hoàn thiện, xác nhận diện tích, kích thước, chiều cao thông thủy các phòng kỹ thuật như phòng trạm biến áp, phòng máy phát (kể cả yêu cầu gác lửng phòng trạm biến áp); vị trí, diện tích phòng điện yếu, buồng phân phối từng tầng; vị trí, kích thước hộp kỹ thuật điện mạnh chính; yêu cầu chiều dày sàn; vị trí, diện tích giếng gió vào và gió ra làm mát máy phát;", "th": "ไฟฟ้าส่งให้สถาปัตยกรรม: เพิ่มเติม ปรับปรุง และยืนยันพื้นที่ ขนาด และความสูงโปร่งของห้องเครื่อง เช่น ห้องหม้อแปลงและจ่ายไฟฟ้า และห้องเครื่องกำเนิดไฟฟ้า (รวมข้อกำหนดชั้นลอยของห้องหม้อแปลง) ตำแหน่งและพื้นที่ห้องไฟฟ้าอ่อนและห้องจ่ายไฟย่อยแต่ละชั้น ตำแหน่งและขนาดช่องไฟฟ้าแรงหลัก ข้อกำหนดความหนาของพื้น และตำแหน่งและพื้นที่ช่องอากาศเข้าและออกเพื่อระบายความร้อนเครื่องกำเนิดไฟฟ้า"},
    "智能化提建筑：补充完善确认进线间、网络机房、电话交换机房、有线电视机房、手机信号覆盖机房、安防控制室、卫星电视等机房面积尺寸净高及特殊机房架空地板要求，弱电竖井大型会议室控制室位置尺寸，楼板厚度要求": {"en": "intelligent systems to architectural: supplement, improve and confirm the area, dimensions and clear height of machine rooms such as entrance rooms, network machine rooms, telephone exchange rooms, cable TV rooms, mobile signal coverage rooms, security control rooms and satellite TV rooms, as well as the raised floor requirements of special machine rooms, the positions and dimensions of weak-current shafts, large conference rooms and control rooms, and floor thickness requirements", "ja": "智能化提建築：引込線室、ネットワーク機室、電話交換機室、有線テレビ機室、携帯信号カバレッジ機室、安防制御室、衛星テレビなどの機械室面積・寸法・天井高及び特殊機械室のOA床要求、弱電シャフト・大型会議室・制御室の位置・寸法、床版厚さ要求を補充・改善・確認する", "ko": "지능화가 건축에: 인입실, 네트워크 기계실, 전화교환기실, 유선TV 기계실, 휴대폰 신호 커버리지 기계실, 안보 통제실, 위성TV 등의 기계실 면적·치수·순높이 및 특수 기계실 액세스 플로어 요구, 약전 샤프트·대형 회의실·통제실 위치·치수, 바닥 슬래브 두께 요구를 보완·개선·확인합니다", "fr": "systèmes intelligents à architecture : compléter, améliorer et confirmer les surfaces, dimensions et hauteurs nettes des locaux techniques tels que les locaux d'introduction, les locaux de réseaux, les locaux de centraux téléphoniques, les locaux de télévision par câble, les locaux de couverture de signal mobile, les salles de contrôle de sécurité et les locaux de télévision par satellite, ainsi que les exigences de plancher surélevé des locaux techniques spéciaux, les positions et dimensions des gaines courants faibles, des grandes salles de réunion et des salles de contrôle, et les exigences d'épaisseur des planchers", "de": "Gebäudeautomation an Architektur: die Flächen, Abmessungen und lichten Höhen von Technikräumen wie Eingangsräumen, Netzwerktechnikräumen, Telefonvermittlungsräumen, Kabelfernsehräumen, Räumen für Mobilfunksignalversorgung, Sicherheitskontrollräumen und Satellitenfernsehräumen sowie die Anforderungen an Doppelböden spezieller Technikräume, die Positionen und Abmessungen von Schwachstromschächten, großen Konferenzräumen und Kontrollräumen und die Anforderungen an die Geschossdeckenstärke", "es": "sistemas inteligentes a arquitectura: complementar, mejorar y confirmar las superficies, dimensiones y alturas libres de los locales técnicos como salas de entrada, salas de redes, salas de centralitas telefónicas, salas de televisión por cable, salas de cobertura de señal móvil, salas de control de seguridad y salas de televisión por satélite, así como los requisitos de suelo técnico elevado de los locales técnicos especiales, las posiciones y dimensiones de los patinillos de corrientes débiles, de las grandes salas de reuniones y de las salas de control, y los requisitos de espesor de forjados", "pt": "sistemas inteligentes para arquitetura: complementar, aperfeiçoar e confirmar as áreas, dimensões e pés-direitos dos ambientes técnicos como salas de entrada, salas de redes, salas de centrais telefônicas, salas de TV a cabo, salas de cobertura de sinal móvel, salas de controle de segurança e salas de TV por satélite, bem como os requisitos de piso elevado dos ambientes técnicos especiais, as posições e dimensões dos shafts de corrente fraca, grandes salas de reunião e salas de controle, e os requisitos de espessura de lajes", "ru": "интеллектуальные системы — архитектуре: дополнить, уточнить и подтвердить площади, размеры и высоты помещений (вводных, серверных, АТС, кабельного телевидения, сотовой связи, диспетчерской охраны, спутникового телевидения и т.п.), а также требования к фальшполам специальных помещений, положение и размеры слаботочных шахт, больших конференц-залов и диспетчерских, требования к толщине перекрытий", "ar": "والأنظمة الذكية إلى العمارة: استكمال وتحسين وتأكيد مساحات وأبعاد وارتفاعات غرف الآلات مثل غرف الإدخال وغرف الشبكات وغرف مقاسم الهاتف وغرف التلفزيون الكبلي وغرف تغطية إشارة الهاتف المحمول وغرف التحكم الأمنية وغرف التلفزيون الفضائي، بالإضافة إلى متطلبات الأرضيات المرتفعة للغرف الفنية الخاصة، ومواقع وأبعاد أعمدة التيار الضعيف وقاعات الاجتماعات الكبيرة وغرف التحكم، ومتطلبات سماكة الألواح", "it": "sistemi intelligenti a architettura: integrare, migliorare e confermare le aree, le dimensioni e le altezze utili dei locali tecnici come locali ingresso linea, locali rete, locali centralino telefonico, locali TV via cavo, locali copertura segnale mobile, sale controllo sicurezza e locali TV satellitare, nonché i requisiti di pavimento sopraelevato dei locali tecnici speciali, le posizioni e dimensioni dei cavedi debole corrente, delle grandi sale riunioni e delle sale controllo, e i requisiti di spessore dei solai", "vi": "hệ thống thông minh gửi kiến trúc: bổ sung, hoàn thiện, xác nhận diện tích, kích thước, chiều cao thông thủy các phòng kỹ thuật như phòng đầu vào, phòng mạng, phòng tổng đài điện thoại, phòng truyền hình cáp, phòng phủ sóng di động, phòng điều khiển an ninh, phòng truyền hình vệ tinh, cùng yêu cầu sàn nâng cho phòng kỹ thuật đặc biệt; vị trí, kích thước hộp kỹ thuật điện yếu, phòng họp lớn, phòng điều khiển; yêu cầu chiều dày sàn", "th": "ระบบอัจฉริยะส่งให้สถาปัตยกรรม: เพิ่มเติม ปรับปรุง และยืนยันพื้นที่ ขนาด และความสูงโปร่งของห้องเครื่อง เช่น ห้องอินพุต ห้องเครือข่าย ห้องตู้สลับโทรศัพท์ ห้องทีวีเคเบิล ห้องครอบคลุมสัญญาณโทรศัพท์มือถือ ห้องควบคุมความปลอดภัย และห้องทีวีดาวเทียม รวมถึงข้อกำหนดพื้นยกสูงของห้องเครื่องพิเศษ ตำแหน่งและขนาดช่องไฟฟ้าอ่อน ห้องประชุมขนาดใหญ่ และห้องควบคุม และข้อกำหนดความหนาของพื้น"},
}

# 简洁版节点清单"内容详专业配合规定第X节Y"引用句式的 12 语言模板译文
# {sec} 会被替换为完整章节编号（如 2.2.1）。
SECTION_REF_TRANS = {
    "en": "For content details, see Section {sec} of the Professional Coordination Regulations for Architectural Engineering Design Projects",
    "ja": "内容の詳細は『建築工程設計項目専門配合規定』第{sec}節を参照",
    "ko": "내용 상세는 《건축공정설계 전문배합규정》 제{sec}절 참조",
    "fr": "Pour le détail du contenu, voir l'article {sec} du Règlement de coordination professionnelle des projets de conception d'ingénierie architecturale",
    "de": "Einzelheiten zum Inhalt siehe Abschnitt {sec} der Fachkoordinationsregelung für architektonische Ingenieurentwurfsprojekte",
    "es": "Para los detalles del contenido, véase el apartado {sec} del Reglamento de coordinación profesional de proyectos de diseño de ingeniería arquitectónica",
    "pt": "Para detalhes do conteúdo, ver o artigo {sec} do Regulamento de coordenação profissional de projetos de engenharia arquitetônica",
    "ru": "Подробности содержания см. в разделе {sec} Правил профессиональной координации проектов архитектурного проектирования",
    "ar": "لتفاصيل المحتوى، يُرجى الرجوع إلى القسم {sec} من لوائح التنسيق المهني لمشاريع التصميم الهندسي المعماري",
    "it": "Per i dettagli del contenuto, si veda la sezione {sec} del Regolamento di coordinamento professionale dei progetti di ingegneria architettonica",
    "vi": "Chi tiết nội dung xem Phần {sec} của Quy định phối hợp chuyên môn các dự án thiết kế công trình kiến trúc",
    "th": "ดูรายละเอียดเนื้อหาได้ที่ข้อ {sec} ของข้อกำหนดการประสานงานวิชาชีพโครงการออกแบบวิศวกรรมสถาปัตยกรรม",
}


def translate_section_ref_cn(section, languages):
    """翻译简洁版节点清单的章节引用句式"内容详专业配合规定第X节Y"。

    section 为完整章节编号（如 2.2.1）；languages 为空时保持中文句式。
    """
    if not section:
        return ""
    cn_text = "内容详专业配合规定第{0}节{1}".format(section.split(".")[0], section)
    if not languages:
        return cn_text
    parts = [cn_text]
    for code in languages:
        tmpl = SECTION_REF_TRANS.get(code)
        if tmpl:
            parts.append("　[{0}] {1}".format(code.upper(), tmpl.format(sec=section)))
    if len(parts) == 1:
        return cn_text
    return "".join(parts)


def translate_content_cn(text, languages):
    """在节点"具体内容/事项"中文文本后追加所选语言译文。

    优先命中 CONTENT_TRANS 条目级映射（基于源文档原文的高频/通用条目）；
    未收录的条目保持中文原文，严禁编造或机翻硬译。
    """
    if not text or not languages:
        return text
    entry = CONTENT_TRANS.get(text)
    if not entry:
        return text
    parts = [text]
    for code in languages:
        trans = entry.get(code)
        if trans:
            parts.append("　[{0}] {1}".format(code.upper(), trans))
    if len(parts) == 1:
        return text
    return "".join(parts)


def normalize_languages(langs):
    """规范化语言代码列表：去重、按 LANG_OPTIONS 顺序、过滤未知代码。"""
    if not langs:
        return []
    out = []
    seen = set()
    for item in langs:
        code = str(item).strip().lower()
        if code in LANG_OPTIONS and code not in seen:
            seen.add(code)
            out.append(code)
    return out


def translate_text_cn(text, languages):
    """在中文文本后追加所选语言译文；词典未覆盖的词条保持中文（不编造）。

    多个语言按 LANG_OPTIONS 顺序依次追加，格式：中文　[EN] 译文　[JA] 译文。
    """
    if not text or not languages:
        return text
    entry = LANG_DICT.get(text)
    if not entry:
        return text
    parts = [text]
    for code in languages:
        trans = entry.get(code)
        if trans:
            parts.append("　[{0}] {1}".format(code.upper(), trans))
    if len(parts) == 1:
        return text
    return "".join(parts)

# v4.0：4 种输出风格配色（桌面与安卓共用；风格名保持英文小写）
# 硬性偏好各风格统一保留：休息日淡绿 #E2EFDA、节点重合浅红 #FFC7CE、斑马纹。
STYLE_PALETTES = {
    "classic": {  # 经典黑白打印友好：深蓝表头 + 浅蓝标题字 + 浅蓝斑马纹
        "desc": "经典黑白（打印友好）：深蓝表头、白字，浅蓝斑马纹",
        "header": "2F5496", "header_text": "FFFFFF", "title": "1F3864",
        "zebra": "EAF1F8", "border": "808080", "accent": "C00000",
        "week_col": "D9E2F3",
    },
    "blue": {  # 蓝白商务：湛蓝表头 + 深蓝标题 + 淡蓝斑马纹
        "desc": "蓝白商务：湛蓝表头、白字，淡蓝斑马纹",
        "header": "1F6FB2", "header_text": "FFFFFF", "title": "0B3B66",
        "zebra": "DDEBF7", "border": "9DC3E6", "accent": "B45F06",
        "week_col": "BDD7EE",
    },
    "green": {  # 淡绿清新：墨绿表头 + 青绿标题 + 浅绿斑马纹（斑马纹与休息日淡绿区分）
        "desc": "淡绿清新：墨绿表头、白字，浅绿斑马纹",
        "header": "3A7A3D", "header_text": "FFFFFF", "title": "1F5C2E",
        "zebra": "E8F5EC", "border": "A9C8A0", "accent": "9C0006",
        "week_col": "C6E0B4",
    },
    "warm": {  # 暖色简约：砖红表头 + 深褐标题 + 浅橙斑马纹
        "desc": "暖色简约：砖红表头、白字，浅橙斑马纹",
        "header": "B4532A", "header_text": "FFFFFF", "title": "7B3F00",
        "zebra": "FBE5D6", "border": "D9B08C", "accent": "B4532A",
        "week_col": "F4CCCC",
    },
}

# =====================================================================
# v4.1：专业作业清单（桌面与安卓共用）
# =====================================================================
# 底色：发图（提出）行淡橙 #FCE4D6，收图（接收）行淡蓝 #DDEBF7。
WORK_LIST_SEND_FILL = "FCE4D6"   # 发图（提出）
WORK_LIST_RECV_FILL = "DDEBF7"   # 收图（接收）

# 各设计阶段专业范围：初步设计 8 专业（含经济），施工图 7 专业（去经济）。
WORK_LIST_PROFESSIONS = {
    "preliminary": ["建筑", "总平面", "结构", "给排水", "暖通", "电气", "智能化", "经济"],
    "construction": ["建筑", "总平面", "结构", "给排水", "暖通", "电气", "智能化"],
}

# 内部专业名（节点数据/文档用语）-> 标准专业名映射；"机电"为水、暖、电、智能化统称。
PROF_ALIASES = {
    "建筑": "建筑", "总图": "总平面", "结构": "结构",
    "水": "给排水", "暖": "暖通", "电": "电气",
    "智能化": "智能化", "经济": "经济",
    "机电": None,  # 在 _split_prof_tokens 中展开
}


def _norm_prof_token(tok):
    """规范化单个专业 token：去“给/专业”等前后缀并映射到标准专业名。"""
    t = tok.strip().rstrip("专业").strip()
    if t == "给排水":
        return ["给排水"]
    t = t.lstrip("给").strip()
    if t == "机电":
        return ["给排水", "暖通", "电气", "智能化"]
    std = PROF_ALIASES.get(t)
    if std:
        return [std]
    # 无法识别时保持原名（不编造，仅用于收图/发图归类）
    return [t]


def _split_prof_tokens(s):
    """把“水、暖、电、智能化”之类的组合拆成标准专业名列表。"""
    out = []
    if not s:
        return out
    for seg in s.replace("和", "、").replace("与", "、").split("、"):
        for std in _norm_prof_token(seg):
            if std and std not in out:
                out.append(std)
    return out


def _extract_work_entries(node_map, stage_key, extra_allowed=None, extra_entries=None):
    """从 node_map 按专业标签分组推导专业作业清单条目。

    返回 {标准专业名: [(方向, 日期, 阶段, 标题, 内容), ...]}，方向为
    "发图"（提出）或 "收图"（接收）；内容均摘自节点原文，不新增编造。
    规则：
      * 段落级：内容中“X提Y：…；X提Y：…”按段落拆分，X 为提出者（发图），
        Y 为接收者（收图），段落文本摘自原文；
      * 节点级：标题/内容“XX专业提出/反提资料/提清资料/反馈意见/确认与
        提清资料/专业提…资料/建筑/总图作业图等”识别提出者；“接受专业/
        接收专业：A、B、C”识别接收者列表。
      * v5.1：extra_entries（覆写节点“发图专业/收图专业”列生成的条目）一并
        并入，统一走专业过滤与去重。
    """
    import re
    result = {}
    skip_titles = ("首次专业协调会", "第二次专业协调会", "第三次专业协调会",
                   "管道综合与对图", "管道综合与对图协调会",
                   "专业间核对、会审与会签", "图纸归档", "开始出图", "交付")
    for d in sorted(node_map.keys()):
        for (title, content, stage, section, percent) in node_map[d]:
            if any(k in title for k in skip_titles):
                continue
            # ---- 段落级：X提Y：…
            seg_re = re.compile(
                r"([\u4e00-\u9fa5、]+?)(?:专业)?提(?:(?:给)|(?=[\u4e00-\u9fa5、]+?[：:]))"
                r"([\u4e00-\u9fa5、]+?)[：:]")
            segs = list(seg_re.finditer(content))
            if segs:
                for i, m in enumerate(segs):
                    seg_text = content[m.start():segs[i + 1].start() if i + 1 < len(segs) else len(content)].strip().rstrip("；;")
                    who = _split_prof_tokens(m.group(1))
                    to = _split_prof_tokens(m.group(2))
                    for w in who:
                        result.setdefault(w, []).append(("发图", d, stage, title, seg_text))
                    for t in to:
                        result.setdefault(t, []).append(("收图", d, stage, title, seg_text))
                continue
            # ---- 节点级：提出者
            senders = []
            m = re.match(r"^([\u4e00-\u9fa5、]+?)专业(?:提出|反提资料|提清资料|反馈意见|确认与提清资料)", title)
            if m:
                senders = _split_prof_tokens(m.group(1))
            if not senders:
                m = re.match(r"^([\u4e00-\u9fa5、]+?)专业对([\u4e00-\u9fa5、]+?)所提资料", title)
                if m:
                    senders = _split_prof_tokens(m.group(1))
                    receivers_t = _split_prof_tokens(m.group(2))
                    for r in receivers_t:
                        result.setdefault(r, []).append(("收图", d, stage, title, content))
            if not senders:
                m = re.match(r"^([\u4e00-\u9fa5、]+?)专业提([\u4e00-\u9fa5、]+?)资料", title)
                if m:
                    senders = _split_prof_tokens(m.group(1))
                    receivers_t = _split_prof_tokens(m.group(2))
                    for r in receivers_t:
                        result.setdefault(r, []).append(("收图", d, stage, title, content))
            if not senders and title.startswith("建筑资料的局部变更"):
                senders = ["建筑"]
            if not senders and title.startswith("建筑作业底图"):
                senders = ["建筑"]
            if not senders:
                for key, who in (("建筑方案图", "建筑"), ("建筑作业图", "建筑"),
                                 ("建筑初步设计图", "建筑"), ("建筑资料", "建筑"),
                                 ("建筑第二版作业图", "建筑"), ("总图方案图", "总图"),
                                 ("总图作业图", "总图"), ("总图初步设计图", "总图"),
                                 ("总图第一版作业图", "总图"), ("总图第二版作业图", "总图")):
                    if title.startswith(key):
                        senders = [who]
                        break
            if not senders:
                m = re.search(r"([\u4e00-\u9fa5、]+?)专业提出", content)
                if m:
                    senders = _split_prof_tokens(m.group(1))
            if senders:
                for s in senders:
                    result.setdefault(s, []).append(("发图", d, stage, title, content))
            # ---- 节点级：接收者（接受专业/接收专业列表）
            m = re.search(r"(?:接受专业|接收专业)[：:]\s*([^。；;]+)", content)
            if m:
                for r in _split_prof_tokens(m.group(1)):
                    result.setdefault(r, []).append(("收图", d, stage, title, content))
    # v5.1：并入覆写节点“发图专业/收图专业”列生成的收发图条目
    if extra_entries:
        for prof, lst in extra_entries.items():
            if not prof or not lst:
                continue
            result.setdefault(prof, []).extend(lst)
    # 仅保留本阶段专业；去重（同专业/同日期/同标题/同方向/同内容只留一条）
    allowed = set(WORK_LIST_PROFESSIONS.get(stage_key, [])) | set(extra_allowed or [])
    out = {}
    for prof, entries in result.items():
        if prof not in allowed:
            continue
        seen = set()
        clean = []
        for e in entries:
            key = (e[0], e[1], e[2], e[3], e[4][:40])
            if key in seen:
                continue
            seen.add(key)
            clean.append(e)
        clean.sort(key=lambda x: (x[1], x[2], x[3]))
        out[prof] = clean
    return out


# ================= v5.3.3(U1) 模板样式抽取与界面化 =================

def load_style_palettes():
    """返回内置样式 + config.json 中用户自定义样式的合并表（v5.3.3(U1)）。"""
    _HEX6 = re.compile(r"^[0-9A-Fa-f]{6}$")
    _REQ_COLOR = ("header", "header_text", "title", "zebra", "border", "accent", "week_col")

    def _valid_pal(d):
        if not isinstance(d, dict):
            return False
        for k in _REQ_COLOR:
            if not (isinstance(d.get(k), str) and _HEX6.match(d[k])):
                return False
        return True

    pals = dict(STYLE_PALETTES)
    try:
        cfg_styles = load_config().get("custom_styles") or {}
        for k, v in cfg_styles.items():
            if _valid_pal(v):
                pals[str(k)] = dict(v)
    except Exception:
        pass
    return pals


def load_config():
    """读取 config.json；不存在或损坏时返回默认配置。

    v5.3.3(U2)：config.json 增加 schema_version 版本号；旧版配置自动补齐
    缺省字段并升级版本号；文件损坏时自动备份为 .bak_时间戳 并恢复默认。
    """
    cfg = dict(DEFAULT_CONFIG)
    data = None
    _loaded = False
    _corrupt = False
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                _loaded = True
                for k in DEFAULT_CONFIG:
                    if k in data:
                        cfg[k] = data[k]
                cfg["schema_version"] = CONFIG_SCHEMA_VERSION
            else:
                _corrupt = True
    except Exception:
        _corrupt = True
    if _corrupt:
        _backup_corrupt_config()
        cfg = dict(DEFAULT_CONFIG)
        cfg["schema_version"] = CONFIG_SCHEMA_VERSION
    if _loaded and not isinstance(data.get("custom_styles"), dict):
        cfg["custom_styles"] = {}
    return cfg


def _backup_corrupt_config():
    """v5.3.3(U2)：config.json 损坏时备份为 .bak_YYYYmmdd_HHMMSS。"""
    try:
        if os.path.exists(CONFIG_PATH):
            import time as _t
            ts = _t.strftime("%Y%m%d_%H%M%S")
            bak = CONFIG_PATH + ".bak_" + ts
            with open(CONFIG_PATH, "rb") as f_src, open(bak, "wb") as f_dst:
                f_dst.write(f_src.read())
    except Exception:
        pass


def save_config(cfg):
    """将本次使用的配置写回 config.json（目录不可写时静默跳过）。"""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# =====================================================================
# 一、专业配合节点数据（摘自《建筑工程设计项目专业配合规定》2015版原文）
# =====================================================================
# 数据结构：(时间百分比, 所属阶段, 节点名称, 具体配合内容)
# 百分比为文档规定的“提出时间”（按设计周期百分比计算）；0 表示设计开始日；
# None 表示文档未给出具体百分比（如会审会签/归档，注明在归档/交付前完成），
# 程序将落在设计结束日。percent=100 表示设计结束日。

PRELIMINARY_NODES = [
    # ---- 第一阶段：初设前期阶段（时间周期：初步设计开始日～15%P）----
    (0, "第一阶段 初设前期阶段", "首次专业协调会",
     "项目经理组织、设计主持人主持、各专业负责人参加：介绍项目基本情况，发放《设计任务书》和《设计策划表》；发放经甲方确认的方案图；研究和讨论甲方及政府部门要求，制定统一设计要求和措施；制定本工程《专业配合进度表》并签字生效"),
    (0, "第一阶段 初设前期阶段", "建筑方案图",
     "建筑专业提出，用于各专业初步设计准备图。内容：方案设计说明；全部方案图纸（平面图、立面图、剖面图，深度满足建设部对方案设计深度的要求）；政府、甲方或业主对方案的修改和（或）审批意见；其他需要提出的资料。接受专业：总图、结构、水、暖、电、智能化、经济"),
    (9, "第一阶段 初设前期阶段", "结构专业反馈意见",
     "结构专业提给建筑专业：考虑结构方案、研究建筑方案是否满足结构可行性；提出地质勘察要求；提出结构构造措施（伸缩缝、沉降缝、防震缝位置要求及特殊部位构造做法）；提出人工地基处理方法、处理深度；提出结构选型（砖混、框架、框剪、筒体、转换结构等）；提出基础形式（独立、条形、十字交叉梁、筏基、桩基、箱型等）及基础截面估算尺寸；对建筑方案修改的其他建议"),
    (9, "第一阶段 初设前期阶段", "机电专业反馈意见",
     "机电专业提给建筑专业：考虑机电系统设计方案、研究建筑方案是否满足机电系统可行性；提出机电系统能源站房大致面积、建议位置或区域；提出各分散机房大致面积、位置的确定原则；对主要大型管道井提出设置原则和面积要求；大型设备吊装孔位置和尺寸要求及运输通道要求；对建筑方案修改的其他建议"),
    (15, "第一阶段 初设前期阶段", "总图方案图",
     "总图专业提出。内容：场地总平面（场地红线及红线内已有建筑、构筑物、道路、绿化等的名称及布置情况）；简要市政条件（可采用文字说明）；用地红线；其他需要提出的资料。接收专业：建筑、结构、水、暖、电、智能化、经济"),

    # ---- 第二阶段：第一版作业图阶段（时间周期：27%P～55%P）----
    (27, "第二阶段 第一版作业图阶段", "第二次专业协调会",
     "项目经理组织、设计主持人主持、各专业负责人参加（第二阶段开始日）：第一阶段情况总结；研究和讨论存在的问题，介绍各专业反馈意见的协调情况和结果；发放第一版建筑作业图"),
    (27, "第二阶段 第一版作业图阶段", "建筑作业图（第一版）",
     "建筑专业提出，作为各专业初步设计工作基础用图。内容（在第一阶段基础上）：各层建筑平面图、主要立面图及剖面图（含初步防火分区示意、初步人防平面及口部布置、承重结构位置轴线与尺寸、机电能源站房及分散机房位置、主要建筑构配件、主要建筑设备位置、房间名称与设计标高、二次装修区域范围、主要管道井配电间弱电间位置、屋顶设备与屋面绿化说明等）；文字说明（楼地面墙体屋面及地下室底板初步构造做法、不同功能区域建筑面积和（或）设计使用人数）。接受专业：总图、结构、水、暖、电、智能化"),
    (30, "第二阶段 第一版作业图阶段", "水、暖、电、智能化专业反提资料1",
     "水提暖：特殊房间温湿度参数和通风要求（如水箱间温度要求、气体消防房间名称等）；暖提水：不采暖房间名称；电提暖：特殊房间温湿度参数和通风要求、电气发热量较大的房间设备发热量（或估算电气安装容量，如变配电室、大型计算机主机房）；智能化提暖：各弱电机房、安防控制室温湿度参数和通风要求；智能化提电：各弱电机房、安防控制室、弱电间用电量"),
    (37, "第二阶段 第一版作业图阶段", "水、暖、电、智能化专业反提资料2",
     "水提建筑：补充完善确认水泵房、水处理机房、热交换站、热水机房、水箱间、报警阀室等机房面积尺寸净高要求，集中管道井位置尺寸，水池尺寸与平面定位、地下室集水坑位置尺寸，影响室内布置和外立面的设备（消火栓、冷却塔等）初步位置，建筑最底层垫层厚度要求；水提暖：由动力专业提供热源时提出估算总热负荷和二次水温度参数（热媒为蒸汽时提出凝结水回水方式）；暖提建筑：补充完善确认冷冻机房、水泵房、锅炉房、热交换站、燃气调压站（室）等机房面积尺寸净高，各层分散空调通风机房位置面积，集中管道井位置尺寸，锅炉房等泄爆井（泄爆面）初步位置面积；暖提水：主要用水点房间名称和估算用水量、主要排水点房间名称和排水性质、冷却水总流量进出水温要求和拟定冷水机组台数；暖提电：需防雷接地和防静电的设备名称（油气锅炉房、烟囱、燃气放散管、可燃气体管道等）；电提建筑：补充完善确认变配电所、发电机房等机房面积尺寸净高，各层弱电机房配电小室位置面积，主要强电竖井位置尺寸，楼板厚度要求，发电机冷却进风与排风竖井位置面积；电提水：特殊电气用房消防要求；智能化提建筑：补充完善确认进线间、网络机房、电话交换机房、有线电视机房、手机信号覆盖机房、安防控制室、卫星电视等机房面积尺寸净高，弱电竖井大型会议室控制室位置尺寸，楼板厚度要求"),
    (43, "第二阶段 第一版作业图阶段", "水、暖、电专业反提资料3",
     "水、暖、电提结构：位于楼板上的机房设备和冷却塔的结构荷载；水、暖提电：各强电设备的位置及其电气安装容量（消防设备和平时使用的设备分别提出）"),
    (49, "第二阶段 第一版作业图阶段", "水、暖、电、智能化专业反提资料4",
     "水、暖、电、智能化提结构：最大长边尺寸或直径≥1000mm的承重结构墙和楼板的预留洞初步位置；水、暖、电提总图：各市政管线（给水、排水、热力、燃气、电力、电信、有线等）在本楼的接口位置要求，主要室外管线布置及主要相关构筑物（化粪池、隔油池、水表井、水泵结合器、阀门井、管线检修井、烟囱、室外箱式变压器等）初步位置，有燃油系统时提出室外储油罐占用面积、初步位置和相关技术要求"),
    (55, "第二阶段 第一版作业图阶段", "结构专业反提资料",
     "结构专业提出：楼面、屋面结构梁板平面图；与建筑有联系的挡土墙结构形式（提建筑、总图专业）；设计配合所必需的特殊结构构造做法或详图（提建筑、总图专业）。接受专业：建筑、水、暖、电、经济"),

    # ---- 第三阶段：第二版作业图阶段（时间周期：60%P～67%P）----
    (60, "第三阶段 第二版作业图阶段", "第三次专业协调会",
     "项目经理组织、设计主持人主持、各专业负责人参加（第三阶段开始日）：第二阶段情况总结；检查《专业配合进度表》执行情况，针对各专业配合资料进行协调，解决上阶段配合过程中的问题；安排下一阶段工作"),
    (60, "第三阶段 第二版作业图阶段", "建筑第二版作业图",
     "建筑专业提出完整的电子版文件，作为各专业初步设计正式绘图用底图。内容（在第一版建筑作业图基础上）：全面完整反映第二阶段各专业提资料、协调配合成果；完善防火分区图、人防平面及口部布置与做法、全部房间名称、建筑专业本身尺寸标注；对特定房间提供必要详图放大图（客房、住宅单元放大图、屋面及外墙构造详图）；外门、外窗尺寸及性能要求（光学、热工性能）；对预留竖井标注使用专业。接受专业：总图、结构、水、暖、电、智能化、经济"),
    (65, "第三阶段 第二版作业图阶段", "总图作业图",
     "总图专业提出，作为各专业初步设计正式绘图用底图。内容（在第一阶段基础上）：含规划用地坐标图在内的各建筑单体轴线或外墙定位图；各单体建筑±0.000标高与绝对标高的关系；红线内已有建筑、构筑物、道路、绿化等布置情况及名称；市政管线走向及平面位置；其他影响设计的资料（排洪沟、保留市政管线等）。接受专业：建筑、结构、水、暖、电、经济"),
    (67, "第三阶段 第二版作业图阶段", "机电专业确认与提清资料",
     "此时应提清所有相关的配合资料：水专业向电专业提出各层消火栓、水流指示器的数量；暖专业向电专业提出各层防火阀、排烟阀、加压送风阀的数量；完善、补充和确认第二阶段向各相关专业提出的资料；水、电专业向智能化专业提出需远程计量各类表具的数量、污水池污水泵的数量；暖专业向智能化专业提出空调通风设备自动化控制要求"),

    # ---- 第四阶段：管道综合与对图阶段（时间周期：78%P～87%P）----
    (78, "第四阶段 管道综合与对图阶段", "建筑资料的局部变更",
     "在第二版建筑作业图和机电专业确认与提清资料并协调后的基础上、由于管道综合后发生的局部变更（重大变更应在第二版作业图中完成），采用纸质文件形式将局部变更处提交给相关专业"),
    (78, "第四阶段 管道综合与对图阶段", "管道综合与对图",
     "项目设计主持人组织、各专业负责人参加（必要时设计人参加）。工作内容：验证与核对房间名称和使用功能；验证与核对防火分区；对主要管道进行综合与协调以保证不给施工图设计留下后患；验证与核对各专业所提资料和配合过程中的协商内容是否在相关图纸中得到反映；验证是否满足甲方和政府主管部门要求。管道综合后预留洞位置和尺寸若发生变化，各专业应立即书面通知结构专业（本阶段 78%P～87%P 时间周期内完成）"),
    (87, "第四阶段 管道综合与对图阶段", "机电专业对总图所提资料的确认与局部变更",
     "针对第二阶段水、暖、电专业对总图所提资料（各市政管线接口位置、室外管线布置、室外构筑物位置等）进行确认；根据管道综合情况进行局部调整与变更"),

    # ---- 第五阶段：会审与会签（时间：设计文件归档之前）----
    (None, "第五阶段 会审与会签", "专业间核对、会审与会签",
     "在会签之前，各专业应再次进行专业间核对，确保上述各阶段的结论和成果已经得到完整落实；各专业的设计图纸为已经过规定的校、审和修改程序后的图纸（设计文件归档之前完成）"),

    # ---- 第六阶段：图纸归档（时间：设计文件交付之前）----
    (None, "第六阶段 图纸归档", "图纸归档",
     "按照院相关文件的要求进行（设计文件交付之前完成）"),
]

CONSTRUCTION_NODES = [
    # ---- 第一阶段：施工图前期阶段（时间周期：施工图设计开始日～6%W）----
    (0, "第一阶段 施工图前期阶段", "首次专业协调会",
     "项目经理组织、设计主持人主持、各专业负责人参加（施工图设计阶段开始日）：介绍项目基本情况（未经过我院初步设计的项目同时发放《设计任务书》和《设计策划表》）；发放经审批或甲方确认的建筑初步设计图；发放各专业初步设计审批意见；研究讨论甲方及政府部门要求，制定统一设计要求和措施；各专业研究需要调研和收集的资料并尽快组织调研；限额设计提出经审批（或甲方同意）的初步设计概算；制定本工程《专业配合进度表》并签字生效"),
    (0, "第一阶段 施工图前期阶段", "建筑初步设计图",
     "建筑专业提出，作为各专业施工图设计准备图。内容：全部初步设计建筑图纸（建筑各层平面图、立面图、剖面图等），平面图立面图剖面图应符合规定2.3.2条要求并满足建设部对初步设计深度的规定；其他需要提出的资料。接受专业：总图、结构、水、暖、电、智能化"),
    (5, "第一阶段 施工图前期阶段", "总图初步设计图",
     "总图专业提出，符合规定2.3.3条要求并满足建设部对初步设计深度的规定。接收专业：建筑、结构、水、暖、电、智能化"),
    (6, "第一阶段 施工图前期阶段", "结构专业反馈意见",
     "结构专业提给建筑专业：考虑结构方案，针对建筑初步设计图研究结构方案的可行性；提出地质勘察要求；提出结构构造措施（伸缩缝、沉降缝、防震缝位置要求及特殊部位构造做法）；提出人工地基处理方法、处理深度；提出结构选型；提出基础形式及基础截面尺寸；对建筑方案修改的其他建议"),
    (6, "第一阶段 施工图前期阶段", "机电专业反馈意见",
     "机电专业提给建筑专业：考虑机电系统设计方案，针对建筑初步设计图研究对机电系统的要求，特别是现有初步设计的机房、管道井位置和面积的可行性，提出修改意见；提出新增加的机房和管道井的大致面积和初步位置；大型设备吊装孔位置和尺寸要求及运输通道要求；对建筑方案修改的其他建议"),

    # ---- 第二阶段：第一版作业图阶段（时间周期：15%W～50%W）----
    (15, "第二阶段 第一版作业图阶段", "第二次专业协调会",
     "项目经理组织、设计主持人主持、各专业负责人参加（第二阶段开始日）：第一阶段情况总结；研究和讨论存在的问题，介绍各专业反馈意见的协调情况和结果；发放第一版建筑作业图"),
    (15, "第二阶段 第一版作业图阶段", "建筑资料1（第一版作业图）",
     "建筑专业提出，作为各专业施工图设计工作基础用图。内容（在第一阶段基础上）：反映初步设计审批意见和第一阶段配合成果；各层建筑主要平面图、立面图及剖面图（人防平面及口部布置、防火分区、机电能源站房及分散机房位置、主要建筑构配件、主要建筑设备位置、房间名称与设计标高、二次装修区域及吊顶净高、主要管道井位置、屋顶设备与屋面绿化位置等）；文字说明（楼地面墙体屋面及地下室底板构造做法、不同功能区域建筑面积和（或）设计使用人数）；必要的初步详图和放大图（客房、住宅单元放大图）；标明外门、外窗尺寸及玻璃初步性能要求（光学、热工性能）；室外照明设计的初步要求。接受专业：总图、结构、水、暖、电、智能化、经济"),
    (17, "第二阶段 第一版作业图阶段", "水、暖、电、智能化反提资料1",
     "水提暖：特殊房间温湿度参数和通风要求（如水箱间温度要求、气体消防房间名称等）；暖提水：不采暖房间名称；暖提电：需要防雷接地和防静电的设备名称（油气锅炉房、烟囱、燃气放散管、可燃气体管道等）；电提水：特殊房间消防要求；电提暖：特殊房间温湿度参数和通风要求；智能化提暖：各弱电机房、安防控制室温湿度参数和通风要求；智能化提电：各弱电机房、安防控制室用电量"),
    (20, "第二阶段 第一版作业图阶段", "建筑专业资料2",
     "建筑专业提出：防火分区图；各朝向窗墙比、建筑体型系数等与建筑热工有关的资料"),
    (20, "第二阶段 第一版作业图阶段", "水、暖、电、智能化反提资料2",
     "电专业提给暖专业：电气发热量较大的房间的设备发热量（或估算的电气安装容量），例如变配电室、大型计算机主机房等"),
    (25, "第二阶段 第一版作业图阶段", "水、暖、电、智能化反提资料3",
     "水专业提给暖专业：由动力专业提供热源时，提出估算的总热负荷和二次水的温度参数；如果热媒为蒸汽，还需要提出凝结水的回水方式"),
    (30, "第二阶段 第一版作业图阶段", "水、暖、电、智能化反提资料4",
     "水提建筑：补充完善确认水泵房、水处理机房、热交换站、热水机房、水箱间、报警阀室等机房面积尺寸净高，集中管道井位置尺寸，水池尺寸与平面定位、地下室集水坑位置尺寸，影响建筑室内布置和外立面的设备（消火栓、冷却塔等）初步位置，最底层垫层厚度要求；暖提建筑：补充完善确认冷冻机房、水泵房、锅炉房、热交换站、燃气调压站（室）等机房面积尺寸净高，各层分散空调通风机房位置面积，集中管道井位置尺寸，锅炉房泄爆井（泄爆面）初步位置面积，地板采暖区域或房间名称及对建筑地面做法厚度要求，室外立面百页窗（新风口、排风口等）初步位置尺寸；暖提水：主要用水点房间名称和估算用水量、主要排水点房间名称和排水性质、冷却水总流量进出水温要求和拟定冷水机组台数；电提建筑：补充完善确认变配电室、发电机房等机房面积尺寸净高（含变配电室夹层要求），各层弱电机房配电小室位置面积，主要强电竖井位置尺寸，楼板厚度要求，发电机冷却进风与排风竖井位置面积；智能化提建筑：补充完善确认进线间、网络机房、电话交换机房、有线电视机房、手机信号覆盖机房、安防控制室、卫星电视等机房面积尺寸净高及特殊机房架空地板要求，弱电竖井大型会议室控制室位置尺寸，楼板厚度要求"),
    (35, "第二阶段 第一版作业图阶段", "水、暖、电、智能化反提资料5",
     "水、暖、电、智能化提结构：位于楼板上的机房设备和冷却塔的结构荷载；最大长边尺寸或直径≥1000mm的承重结构墙和楼板的预留洞初步位置；水、暖提电：各强电设备的位置、电压等级及其电气安装容量（消防设备和平时使用的设备分别提出）"),
    (40, "第二阶段 第一版作业图阶段", "总图第一版作业图",
     "总图专业提出（在第一阶段基础上）：含规划用地坐标图在内的各建筑单体轴线或外墙定位图；各单体建筑±0.000标高与绝对标高的关系；红线内已有建筑、构筑物、道路、绿化等的名称、布置与定位；市政管线走向及平面位置；其他影响设计的资料（排洪沟、保留市政管线等）。接受专业：建筑、结构、水、暖、电、经济"),
    (45, "第二阶段 第一版作业图阶段", "结构专业反提资料",
     "结构专业提出：楼面、屋面结构梁板平面图；与建筑有联系的挡土墙结构形式（提建筑、总图专业）；设计配合所必需的特殊结构构造做法或详图（提建筑、总图专业）。接受专业：建筑、水、暖、电、智能化、经济"),
    (50, "第二阶段 第一版作业图阶段", "水、暖、电、智能化反提资料6",
     "水提总图：各市政管网（给水、排水、热水等）至本楼接口大致位置，主要室外管道布置及主要构筑物（化粪池、隔油池、水表井、水泵接合器等）大致位置；暖提总图：市政管网（热力、燃气）至本楼接口大致位置，主要室外管道布置及构筑物（阀门井等）大致位置；电、智能化提总图：市政管线（电力、通讯、有线电视等）至本楼接口大致位置，主要室外管线布置及构筑物（检修井等）大致位置"),

    # ---- 第三阶段：第二版作业图阶段（时间周期：55%W～65%W）----
    (55, "第三阶段 第二版作业图阶段", "第三次专业协调会",
     "项目经理组织、设计主持人主持、各专业负责人参加（第三阶段开始日）：第二阶段情况总结；检查《专业配合进度表》执行情况，针对各专业配合资料进行协调，解决上阶段配合过程中的问题；安排下一阶段工作（本阶段各专业提资料如与第二阶段相比只是局部修改，可采用“局部纸质资料+完整电子文件”的方式，局部纸质资料应反映全部修改内容并与完整电子文件内容相一致）"),
    (55, "第三阶段 第二版作业图阶段", "建筑资料1（第二版作业图）",
     "建筑专业提出（在第一版建筑作业图基础上）：全面完整反映第二阶段各专业提资料、协调配合成果；完善补充第一版作业图内容（准确的防火分区图、人防平面及口部布置与详细做法、全部房间名称、建筑专业本身的三道尺寸标注等）；对特定房间提供完整详图放大图（核心筒、客房、住宅单元放大图）；对预留竖井标注使用专业（水、暖、强电、弱电）。接受专业：总图、结构、水、暖、电、智能化、经济"),
    (60, "第三阶段 第二版作业图阶段", "总图第二版作业图",
     "总图专业提出（在第二阶段基础上）：全面完整反映第二阶段各专业提资料、协调配合成果；完善补充第一版作业图内容；标注出所有红线区域内的已有和新建建筑、构筑物、道路、绿化等的名称、布置与定位，以及市政管线（管沟）的平面位置尺寸。接受专业：建筑、结构、水、暖、电、智能化、经济"),
    (60, "第三阶段 第二版作业图阶段", "结构专业反提资料",
     "结构专业提出：根据第二阶段中各专业反提资料和建筑第二版作业图，提出补充和调整后的楼面、屋面结构平面图（模板图）；梁、板、柱、墙的截面与尺寸（包括定位尺寸）图；特殊结构的构造做法和详图。接受专业：建筑、总图、水、暖、电、智能化、经济"),
    (60, "第三阶段 第二版作业图阶段", "水、暖、电、智能化反提资料1",
     "水提建筑：补齐全部管道井尺寸和位置要求；水提电：消火栓、水流指示器、报警阀、控制阀等的位置及设备和附件的控制要求；暖提建筑：补齐全部管道井尺寸和位置要求、补齐全部室外立面百页窗（新风口、排风口等）位置和尺寸；暖提电：排烟阀、加压送风阀、需要电气控制的防火阀以及自动控制阀的位置及设备和附件的控制要求；电提建筑：补齐全部管道井尺寸和位置要求；水、暖、电、智能化提结构：最大长边尺寸或直径≥1000mm的承重墙、楼板及梁上预留洞的准确位置"),
    (65, "第三阶段 第二版作业图阶段", "水、暖、电反提资料2",
     "水、暖、电、智能化提结构：吊挂管道和设备的楼板荷载；需要结构专业进行设计的大型设备的基础尺寸、设备重量及相关要求；水、暖专业中需要结构专业进行加固的管道固定支架的位置和推力；水、暖提建筑：排水沟的位置、尺寸、坡度等要求；水专业提建筑：特殊房间（如气体消防房间等）对建筑门、窗等洞口的要求"),
    (70, "第四阶段 管道综合与对图阶段", "建筑资料2",
     "建筑专业提出：楼梯间、卫生间放大详图；屋面及墙身剖面或构造详图、必要的结点放大图；各个门的开启方向；非承重墙中的构造柱尺寸和位置；非承重墙过梁的尺寸和位置"),
    (75, "第四阶段 管道综合与对图阶段", "水、暖专业提建筑、结构资料",
     "需要结构专业配合设计的水池、积水坑等详图（包括外形尺寸、厚度、预留洞、相关做法、保温要求等）"),

    # ---- 第四阶段：管道综合与对图阶段（时间周期：71%W～76%W）----
    (71, "第四阶段 管道综合与对图阶段", "管道综合与对图协调会",
     "设计主持人组织、各专业负责人参加（必要时相关专业设计人参加）：第三阶段情况总结，安排下一阶段工作；检查《专业配合进度表》执行情况，针对各专业配合资料进行协调，解决上阶段配合过程中的问题"),
    (71, "第四阶段 管道综合与对图阶段", "管道综合与对图",
     "设计主持人组织：对所有管道进行综合与协调，防止管道打架；验证与核对房间名称和使用功能；验证与核对防火分区；验证与核对各专业所提资料和配合过程中的协商内容是否在相关图纸中得到反映；验证各专业施工图是否符合初步设计配合原则；核对并解决管线与建筑图中的构造柱、过梁等可能存在的打架情况（本阶段 71%W～76%W 时间周期内完成）"),

    # ---- 第五阶段：提清资料阶段（时间周期：76%W～82%W）----
    (76, "第五阶段 提清资料阶段", "水专业提清资料",
     "水专业提资料（在第三阶段所提资料和管道综合成果基础上）：提建筑——地漏位置和房间坡度、坡向要求，非承重墙预留洞（根据需要提出）；提总图——市政管网（给水、排水、热水等）至本楼接口准确位置，主要室外管道布置及构筑物（化粪池、隔油池、水表井、水泵接合器等）准确位置；提结构——提清全部承重墙、楼板及梁上预留洞准确位置和尺寸，需要结构设计的总图管线固定支架准确位置和推力；提电——完善补充所有本专业控制元器件、控制阀等的位置与控制要求；提智能化——完善补充需远程计量冷热水表具的位置、污水池污水泵的位置（本阶段 76%W～82%W 内完成）"),
    (76, "第五阶段 提清资料阶段", "暖专业提清资料",
     "暖专业提资料（在第三阶段所提资料和管道综合成果基础上）：提建筑——确认和补齐全部室外新风口和排风口的尺寸和位置，非承重墙预留洞（根据需要提出）；提总图——市政管网（热力、燃气）至本楼接口准确位置，主要室外管道布置及构筑物（阀门井等）准确位置；提结构——提清全部承重墙、楼板及梁上预留洞准确位置和尺寸，需要结构设计的总图管线固定支架准确位置和推力；提电——完善补充所有本专业控制元器件、控制阀等的位置与控制要求；提智能化——提出空调通风设备自动化控制要求（本阶段 76%W～82%W 内完成）"),
    (76, "第五阶段 提清资料阶段", "电专业提清资料",
     "电专业提资料（在第三阶段所提资料和管道综合成果基础上）：提建筑——各房间门开启方向的要求，各消防控制按钮的位置要求；提总图——市政管线（电力、通讯等）至本楼接口准确位置，主要室外管线布置及构筑物（检修井等）准确位置；提结构——提清全部承重墙、楼板及梁上预留洞准确位置和尺寸；提智能化——完善补充需远程计量电表的位置，完善和补充需由建筑设备管理系统监视和控制的配电箱位置（本阶段 76%W～82%W 内完成）"),
    (76, "第五阶段 提清资料阶段", "智能化专业提清资料",
     "智能化专业提资料：对电气专业提出平面上数据插座、有线电视插座等需预留~220V电源的位置（本阶段 76%W～82%W 内完成）"),

    # ---- 第六阶段：建筑作业底图阶段（时间要求：85%W）----
    (85, "第六阶段 建筑作业底图阶段", "建筑作业底图",
     "建筑专业提出本版建筑图纸及相关的资料，应反映上述各过程配合后的综合成果，作为各专业最终出图用的建筑底图；本版图纸可以采用电子文件，但与第三阶段（第二版建筑作业图）有修改之处应在电子文件中做出明确标记、标识或书面的文字说明；电子文件按图中信息分层管理；各专业接收后应根据本专业图纸表达特点进行必要修整，删除与本专业图纸表达无关的图形、文字、尺寸等"),

    # ---- 第七阶段：会审与会签（时间：设计文件归档之前）----
    (None, "第七阶段 会审与会签", "专业间核对、会审与会签",
     "在会签之前，各专业应再次进行专业间核对，确保上述各阶段的结论和成果已经得到完整落实；各专业的会审图纸应为已经过规定的校、审和修改程序后的图纸；会审完成后在各专业图纸上进行会签（设计文件归档之前完成）"),

    # ---- 第八阶段：图纸归档（时间：设计文件交付之前）----
    (None, "第八阶段 图纸归档", "图纸归档",
     "按照院相关文件的要求进行（设计文件交付之前完成）"),
]

# 各阶段时间周期（用于表格底部备注，摘自文档原文）
PRELIMINARY_PERIODS = [
    ("第一阶段 初设前期阶段", "初步设计开始日～15%P"),
    ("第二阶段 第一版作业图阶段", "27%P～55%P"),
    ("第三阶段 第二版作业图阶段", "60%P～67%P"),
    ("第四阶段 管道综合与对图阶段", "78%P～87%P"),
    ("第五阶段 会审与会签", "设计文件归档之前"),
    ("第六阶段 图纸归档", "设计文件交付之前"),
]
CONSTRUCTION_PERIODS = [
    ("第一阶段 施工图前期阶段", "施工图设计开始日～6%W"),
    ("第二阶段 第一版作业图阶段", "15%W～50%W"),
    ("第三阶段 第二版作业图阶段", "55%W～65%W"),
    ("第四阶段 管道综合与对图阶段", "71%W～76%W"),
    ("第五阶段 提清资料阶段", "76%W～82%W"),
    ("第六阶段 建筑作业底图阶段", "85%W"),
    ("第七阶段 会审与会签", "设计文件归档之前"),
    ("第八阶段 图纸归档", "设计文件交付之前"),
]

# ---- v3.5 阶段周期表达规则配置（v3.4 引入，v3.5 扩展至第四阶段）----
# 适用条件：阶段标注百分比起止周期（如 76%W～82%W、xx%P～yy%P），且阶段内
#   具体内容无独立精确时间节点（内容仅注明“本阶段 xx%~yy% 内完成”）。
# 处理：阶段内容全部落在阶段周期开始百分比日期格表达；
#   阶段周期结束百分比日期格标注“X阶段结束”。
# 配置项：(阶段名, 开始百分比, 结束百分比, 周期符号, 阶段结束标注文本)
# 核对结论（v3.5 复核）：初步设计 6 个阶段各节点均有独立百分比（第五阶段
#   会审与会签、第六阶段图纸归档为“设计文件归档/交付之前”非百分比标注），
#   不适用本规则；施工图第一阶段（开始日～6%W）、第二阶段（15%W～50%W）、
#   第三阶段（55%W～65%W）各节点均有独立百分比，第六阶段（85%W）为单点
#   标注非起止周期，第七/八阶段为非百分比标注，均不适用；仅第四阶段
#   “管道综合与对图阶段”（71%W～76%W，内容无独立精确时间节点）与第五阶段
#   “提清资料阶段”（76%W～82%W，内容无独立精确时间节点）符合条件。
PERIOD_STAGE_MARKS = {
    "preliminary": [],
    "construction": [
        ("第四阶段 管道综合与对图阶段", 71, 76, "W", "第四阶段（管道综合与对图）结束"),
        ("第五阶段 提清资料阶段", 76, 82, "W", "第五阶段（提清资料）结束"),
    ],
}

NODE_SECTIONS = {
    "preliminary": {
        "首次专业协调会": "2.1.1",
        "建筑方案图": "2.1.2",
        "总图方案图": "2.1.3",
        "结构专业反馈意见": "2.1.4",
        "机电专业反馈意见": "2.1.5",
        "第二次专业协调会": "2.2.1",
        "建筑作业图（第一版）": "2.2.2",
        "水、暖、电、智能化专业反提资料1": "2.2.3",
        "水、暖、电、智能化专业反提资料2": "2.2.4",
        "水、暖、电专业反提资料3": "2.2.5",
        "水、暖、电、智能化专业反提资料4": "2.2.6",
        "结构专业反提资料": "2.2.7",
        "第三次专业协调会": "2.3.1",
        "建筑第二版作业图": "2.3.2",
        "总图作业图": "2.3.3",
        "机电专业确认与提清资料": "2.3.4",
        "建筑资料的局部变更": "2.4.1",
        "管道综合与对图": "2.4.2",
        "机电专业对总图所提资料的确认与局部变更": "2.4.3",
        "专业间核对、会审与会签": "2.5",
        "图纸归档": "2.6",
    },
    "construction": {
        "首次专业协调会": "3.1.1",
        "建筑初步设计图": "3.1.2",
        "总图初步设计图": "3.1.3",
        "结构专业反馈意见": "3.1.4",
        "机电专业反馈意见": "3.1.5",
        "第二次专业协调会": "3.2.1",
        "建筑资料1（第一版作业图）": "3.2.2",
        "水、暖、电、智能化反提资料1": "3.2.3",
        "建筑专业资料2": "3.2.4",
        "水、暖、电、智能化反提资料2": "3.2.5",
        "水、暖、电、智能化反提资料3": "3.2.6",
        "水、暖、电、智能化反提资料4": "3.2.7",
        "水、暖、电、智能化反提资料5": "3.2.8",
        "总图第一版作业图": "3.2.9",
        "结构专业反提资料": "3.2.10",
        "水、暖、电、智能化反提资料6": "3.2.11",
        "第三次专业协调会": "3.3.1",
        "建筑资料1（第二版作业图）": "3.3.2",
        "总图第二版作业图": "3.3.3",
        "结构专业反提资料": "3.3.4",
        "水、暖、电、智能化反提资料1": "3.3.5",
        "水、暖、电反提资料2": "3.3.6",
        "建筑资料2": "3.3.7",
        "水、暖专业提建筑、结构资料": "3.3.8",
        "管道综合与对图协调会": "3.4.1",
        "管道综合与对图": "3.4.2~3.4.3",
        "水专业提清资料": "3.5.1",
        "暖专业提清资料": "3.5.2",
        "电专业提清资料": "3.5.3",
        "智能化专业提清资料": "3.5.4",
        "建筑作业底图": "3.6",
        "专业间核对、会审与会签": "3.7",
        "图纸归档": "3.8",
    },
}

DOC_TITLE = "《建筑工程设计项目专业配合规定》(2015版，中国建筑设计院有限公司，2015年12月01日)"
DOC_SECTION = {
    "preliminary": "第二节　初步设计配合详细要求",
    "construction": "第三节　施工图设计配合详细要求",
}

# =====================================================================
# 二、中国法定节假日数据（放假日期 + 调休补班日期）
# =====================================================================
# 覆盖 2020～2026（官方公布）；2027 起为农历/公历近似值，仅供参考。
# 可通过 --holidays-file 传入 JSON 覆盖：
#   {"holidays": ["2026-10-01", ...], "workdays": ["2026-10-10", ...]}

_HOLIDAYS = {
    2020: ["2020-01-01", "2020-01-24", "2020-01-25", "2020-01-26", "2020-01-27",
           "2020-01-28", "2020-01-29", "2020-01-30", "2020-04-04", "2020-04-05",
           "2020-04-06", "2020-05-01", "2020-05-02", "2020-05-03", "2020-05-04",
           "2020-05-05", "2020-06-25", "2020-06-26", "2020-06-27", "2020-10-01",
           "2020-10-02", "2020-10-03", "2020-10-04", "2020-10-05", "2020-10-06",
           "2020-10-07", "2020-10-08"],
    2021: ["2021-01-01", "2021-01-02", "2021-01-03", "2021-02-11", "2021-02-12",
           "2021-02-13", "2021-02-14", "2021-02-15", "2021-02-16", "2021-02-17",
           "2021-04-03", "2021-04-04", "2021-04-05", "2021-05-01", "2021-05-02",
           "2021-05-03", "2021-05-04", "2021-05-05", "2021-06-12", "2021-06-13",
           "2021-06-14", "2021-09-19", "2021-09-20", "2021-09-21", "2021-10-01",
           "2021-10-02", "2021-10-03", "2021-10-04", "2021-10-05", "2021-10-06",
           "2021-10-07"],
    2022: ["2022-01-01", "2022-01-02", "2022-01-03", "2022-01-31", "2022-02-01",
           "2022-02-02", "2022-02-03", "2022-02-04", "2022-02-05", "2022-02-06",
           "2022-04-03", "2022-04-04", "2022-04-05", "2022-04-30", "2022-05-01",
           "2022-05-02", "2022-05-03", "2022-05-04", "2022-06-03", "2022-06-04",
           "2022-06-05", "2022-09-10", "2022-09-11", "2022-09-12", "2022-10-01",
           "2022-10-02", "2022-10-03", "2022-10-04", "2022-10-05", "2022-10-06",
           "2022-10-07"],
    2023: ["2023-01-01", "2023-01-02", "2023-01-21", "2023-01-22", "2023-01-23",
           "2023-01-24", "2023-01-25", "2023-01-26", "2023-01-27", "2023-04-05",
           "2023-04-29", "2023-04-30", "2023-05-01", "2023-05-02", "2023-05-03",
           "2023-06-22", "2023-06-23", "2023-06-24", "2023-09-29", "2023-09-30",
           "2023-10-01", "2023-10-02", "2023-10-03", "2023-10-04", "2023-10-05",
           "2023-10-06"],
    2024: ["2024-01-01", "2024-02-10", "2024-02-11", "2024-02-12", "2024-02-13",
           "2024-02-14", "2024-02-15", "2024-02-16", "2024-02-17", "2024-04-04",
           "2024-04-05", "2024-04-06", "2024-05-01", "2024-05-02", "2024-05-03",
           "2024-05-04", "2024-05-05", "2024-06-08", "2024-06-09", "2024-06-10",
           "2024-09-15", "2024-09-16", "2024-09-17", "2024-10-01", "2024-10-02",
           "2024-10-03", "2024-10-04", "2024-10-05", "2024-10-06", "2024-10-07"],
    2025: ["2025-01-01", "2025-01-28", "2025-01-29", "2025-01-30", "2025-01-31",
           "2025-02-01", "2025-02-02", "2025-02-03", "2025-02-04", "2025-04-04",
           "2025-04-05", "2025-04-06", "2025-05-01", "2025-05-02", "2025-05-03",
           "2025-05-04", "2025-05-05", "2025-05-31", "2025-06-01", "2025-06-02",
           "2025-10-01", "2025-10-02", "2025-10-03", "2025-10-04", "2025-10-05",
           "2025-10-06", "2025-10-07", "2025-10-08"],
    2026: ["2026-01-01", "2026-01-02", "2026-01-03", "2026-02-15", "2026-02-16",
           "2026-02-17", "2026-02-18", "2026-02-19", "2026-02-20", "2026-02-21",
           "2026-02-22", "2026-02-23", "2026-04-04", "2026-04-05", "2026-04-06",
           "2026-05-01", "2026-05-02", "2026-05-03", "2026-05-04", "2026-05-05",
           "2026-06-19", "2026-06-20", "2026-06-21", "2026-09-25", "2026-09-26",
           "2026-09-27", "2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04",
           "2026-10-05", "2026-10-06", "2026-10-07"],
    # ---- 2027 起为预测值（按天文历法 + 通行放假惯例推算，含 7 天春节与调休补班）----
    # 预测值，以国务院当年公布为准。春节=除夕~初六共7天；元旦/清明/端午/中秋按
    # “与周末拼接3天，周三单放1天”惯例；五一固定5/1~5/5；国庆固定10/1~10/7
    # （2028 中秋10/3并入国庆放8天）；补班日取假期前后最近周末（春节前后各1天）。
    2027: ["2027-01-01", "2027-01-02", "2027-01-03", "2027-02-05", "2027-02-06",
           "2027-02-07", "2027-02-08", "2027-02-09", "2027-02-10", "2027-02-11",
           "2027-04-03", "2027-04-04", "2027-04-05", "2027-05-01", "2027-05-02",
           "2027-05-03", "2027-05-04", "2027-05-05", "2027-06-09", "2027-09-15",
           "2027-10-01", "2027-10-02", "2027-10-03", "2027-10-04", "2027-10-05",
           "2027-10-06", "2027-10-07"],
    2028: ["2028-01-01", "2028-01-02", "2028-01-03", "2028-02-01", "2028-02-02",
           "2028-02-03", "2028-02-04", "2028-02-05", "2028-02-06", "2028-02-07",
           "2028-04-03", "2028-04-04", "2028-04-05", "2028-05-01", "2028-05-02",
           "2028-05-03", "2028-05-04", "2028-05-05", "2028-05-27", "2028-05-28",
           "2028-05-29", "2028-10-01", "2028-10-02", "2028-10-03", "2028-10-04",
           "2028-10-05", "2028-10-06", "2028-10-07", "2028-10-08"],
    2029: ["2028-12-30", "2028-12-31", "2029-01-01", "2029-02-09", "2029-02-10",
           "2029-02-11", "2029-02-12", "2029-02-13", "2029-02-14", "2029-02-15",
           "2029-04-04", "2029-05-01", "2029-05-02", "2029-05-03", "2029-05-04",
           "2029-05-05", "2029-06-15", "2029-06-16", "2029-06-17", "2029-09-21",
           "2029-09-22", "2029-09-23", "2029-10-01", "2029-10-02", "2029-10-03",
           "2029-10-04", "2029-10-05", "2029-10-06", "2029-10-07"],
    2030: ["2030-01-01", "2030-01-02", "2030-01-03", "2030-02-04", "2030-02-05",
           "2030-02-06", "2030-02-07", "2030-02-08", "2030-02-09", "2030-02-10",
           "2030-04-05", "2030-04-06", "2030-04-07", "2030-05-01", "2030-05-02",
           "2030-05-03", "2030-05-04", "2030-05-05", "2030-06-05", "2030-09-11",
           "2030-09-12", "2030-09-13", "2030-10-01", "2030-10-02", "2030-10-03",
           "2030-10-04", "2030-10-05", "2030-10-06", "2030-10-07"],
}

_WORKDAYS = {
    2020: ["2020-01-19", "2020-02-01", "2020-04-26", "2020-05-09", "2020-06-28",
           "2020-09-27", "2020-10-10"],
    2021: ["2021-02-07", "2021-02-20", "2021-04-25", "2021-05-08", "2021-09-18",
           "2021-09-26", "2021-10-09"],
    2022: ["2022-01-29", "2022-01-30", "2022-04-02", "2022-04-24", "2022-05-07",
           "2022-10-08", "2022-10-09"],
    2023: ["2023-01-28", "2023-01-29", "2023-04-23", "2023-05-06", "2023-06-25",
           "2023-10-07", "2023-10-08"],
    2024: ["2024-02-04", "2024-02-18", "2024-04-07", "2024-04-28", "2024-05-11",
           "2024-09-14", "2024-09-29", "2024-10-12"],
    2025: ["2025-01-26", "2025-02-08", "2025-04-27", "2025-09-28", "2025-10-11"],
    2026: ["2026-01-04", "2026-02-14", "2026-02-28", "2026-04-26", "2026-05-09",
           "2026-10-10"],
    # 2027 起为预测值（与 _HOLIDAYS 对应，以国务院当年公布为准）
    2027: ["2027-01-31", "2027-02-13", "2027-04-25", "2027-05-08",
           "2027-09-26", "2027-10-09"],
    2028: ["2028-01-22", "2028-02-11", "2028-04-08", "2028-04-30",
           "2028-05-06", "2028-09-30", "2028-10-14"],
    2029: ["2029-02-04", "2029-02-17", "2029-04-29", "2029-05-06",
           "2029-09-30", "2029-10-13"],
    2030: ["2030-01-05", "2030-01-27", "2030-02-16", "2030-04-28",
           "2030-05-11", "2030-09-14", "2030-09-29", "2030-10-12"],
}


def _holiday_item_to_date(item):
    """v5.3.3(U3)：兼容节假日 JSON 中“YYYY-MM-DD”字符串与 {"date": "YYYY-MM-DD", "name": ...} 条目。"""
    if isinstance(item, dict):
        return date.fromisoformat(str(item.get("date") or ""))
    return date.fromisoformat(str(item))


def load_holidays(holidays_file=None):
    """返回 (假期集合, 调休上班集合)，可被外部 JSON 覆盖/补充。"""
    holidays = set()
    workdays = set()
    for year, days in _HOLIDAYS.items():
        for d in days:
            try:
                holidays.add(date.fromisoformat(d))
            except ValueError:
                pass
    for year, days in _WORKDAYS.items():
        for d in days:
            try:
                workdays.add(date.fromisoformat(d))
            except ValueError:
                pass
    if holidays_file:
        with open(holidays_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "holidays" in data:
            for d in data["holidays"]:
                holidays.add(_holiday_item_to_date(d))
        if "workdays" in data:
            for d in data["workdays"]:
                workdays.add(_holiday_item_to_date(d))
        if "clear_builtin" in data and data["clear_builtin"]:
            holidays = set()
            workdays = set()
            for d in data.get("holidays", []):
                holidays.add(date.fromisoformat(str(d)))
            for d in data.get("workdays", []):
                workdays.add(date.fromisoformat(str(d)))
    return holidays, workdays


def is_rest_day(d, holidays, workdays):
    """休息日 = 周六/周日或法定节假日，扣除调休上班日。"""
    if d in workdays:
        return False
    if d.weekday() >= 5:
        return True
    return d in holidays


# =====================================================================
# 三、里程碑识别
# =====================================================================

_MILESTONE_KEYS = ("协调会", "会审", "会签", "图纸归档", "出图", "交付")


def is_milestone(title):
    """识别里程碑节点：各次专业协调会、专业间核对/会审与会签、图纸归档、出图等。"""
    return any(k in title for k in _MILESTONE_KEYS)


# =====================================================================
# 四、日期换算与表格渲染
# =====================================================================

def percent_to_date(start, end, percent, round_mode="round"):
    """按文档规则将百分比换算为日期：offset = round(percent/100 * 周期天数)。

    v5.3.2(A4)：新增 round_mode 取整模式参数：
      "round"  = 四舍五入（默认，保持 v5.3.1 及以前行为）；
      "floor"  = 向下取整（只减不加）；
      "ceil"   = 向上取整（只加不减）。
    """
    if percent is None:
        return end  # 文档未给百分比（归档/会签前完成）落在结束日
    period = (end - start).days
    raw = percent / 100.0 * period
    if round_mode == "floor":
        offset = int(raw)
    elif round_mode == "ceil":
        offset = int(raw) if raw == int(raw) else int(raw) + 1
    else:
        # 传统四舍五入（避免 round() 银行家舍入在 x.5 时向下取整）
        offset = int(raw + 0.5)
    return start + timedelta(days=offset)


def build_node_map(nodes, start, end, stage_key, round_mode="round"):
    """节点列表 -> {日期: [(节点名, 内容, 阶段, 章节编号, 原始百分比), ...]}

    原始百分比用于重合判定：同一日期仅当存在两个及以上原始时间设定（百分比）
    不同的节点时才判定为异常重合。None 表示文档未给百分比（会审/归档等）。
    v5.3.2(A4)：round_mode 透传 percent_to_date（floor/ceil/round）。
    """
    result = {}
    section_map = NODE_SECTIONS.get(stage_key, {})
    for percent, stage, title, content in nodes:
        d = percent_to_date(start, end, percent, round_mode)
        if d < start or d > end:
            continue
        section = section_map.get(title, "")
        result.setdefault(d, []).append((title, content, stage, section, percent))
    return result


def apply_stage_period_rules(node_map, start, end, stage_key, round_mode="round"):
    """v3.5：阶段周期表达规则（v3.4 引入，v3.5 扩展至第四阶段）。

    对“阶段标注百分比起止周期（如 76%W～82%W、xx%P～yy%P）且阶段内具体内容
    无独立精确时间节点”的阶段：
      1) 将该阶段全部节点内容重定位到阶段周期开始百分比换算日期格；
      2) 在阶段周期结束百分比换算日期格追加“X阶段结束”标注条目
         （percent 标记为 "_stage_end"，不参与重合判定、不计入节点统计）。

    “阶段结束”标注与普通节点一样参与“避开节假日”迁移：落在休息日时迁移到
    最近工作日并注明原日期，保证计划表所有日期标记在工作日表达，语义不丢失。
    “阶段结束”标注所在日期格若与其它节点同格，该格不标注浅红、不触发
    “由于整体周期过短，导致节点重合，需要人工协调”提示（重合判定与单格
    短周期提示均只针对普通节点，"_stage_end" 被完全排除）。
    """
    marks = PERIOD_STAGE_MARKS.get(stage_key, [])
    if not marks:
        return
    for stage_name, start_pct, end_pct, symbol, end_mark_text in marks:
        start_d = percent_to_date(start, end, start_pct, round_mode)
        end_d = percent_to_date(start, end, end_pct, round_mode)
        # 1) 将该阶段所有节点重定位到周期开始日期格
        relocated = []
        for d, items in list(node_map.items()):
            kept = []
            for item in items:
                if item[2] == stage_name and item[4] != "_stage_end":
                    relocated.append((item[0], item[1], item[2], item[3], start_pct))
                else:
                    kept.append(item)
            if kept:
                node_map[d] = kept
            else:
                node_map.pop(d, None)
        for item in relocated:
            node_map.setdefault(start_d, []).append(item)
        # 2) 阶段结束标注
        mark_content = ("该阶段周期为 {0}%{1}～{2}%{1}，阶段内内容无独立精确时间节点，"
                        "已全部在周期开始日期（{3}）表格内表达；此为该阶段周期结束日期格标注。").format(
                            start_pct, symbol, end_pct, start_d.strftime("%Y-%m-%d"))
        node_map.setdefault(end_d, []).append(
            (end_mark_text, mark_content, stage_name, "", "_stage_end"))


def nearest_workday(d, start, end, holidays, workdays):
    """在 [start, end] 范围内寻找距 d 最近的工作日；距离相同时优先未来（顺延）。

    返回工作日日期；范围内找不到时返回 None。
    """
    if not is_rest_day(d, holidays, workdays):
        return d
    max_dist = (end - start).days + 2
    for dist in range(1, max_dist):
        future = d + timedelta(days=dist)
        past = d - timedelta(days=dist)
        if future <= end and not is_rest_day(future, holidays, workdays):
            return future
        if past >= start and not is_rest_day(past, holidays, workdays):
            return past
    return None


def migrate_rest_day_nodes(node_map, start, end, holidays, workdays, skip_titles=()):
    """把落在休息日（节假日/周末）的节点迁移到最近工作日，并注明原日期。

    迁移后节点内容追加一行：由X月X日（节假日）迁移而来。
    skip_titles 中的节点不参与迁移（子串匹配；v3.7 起“交付”节点固定为整体
    周期最终日期，不参与避开节假日迁移；“开始出图”若落休息日则正常前移）。
    """
    entries = []
    for d, items in list(node_map.items()):
        for item in items:
            entries.append((d,) + item)
    node_map.clear()
    for orig_d, title, content, stage, section, percent in entries:
        if any(k in title for k in skip_titles):
            node_map.setdefault(orig_d, []).append(
                (title, content, stage, section, percent))
            continue
        if is_rest_day(orig_d, holidays, workdays):
            nd = nearest_workday(orig_d, start, end, holidays, workdays)
            if nd is not None and nd != orig_d:
                content = "{0}\n（由{1}月{2}日（节假日）迁移而来）".format(
                    content, orig_d.month, orig_d.day)
                d = nd
            else:
                d = orig_d  # 范围内找不到工作日，保持原位
        else:
            d = orig_d
        node_map.setdefault(d, []).append((title, content, stage, section, percent))
    return node_map


def apply_drawing_schedule(node_map, start, end, drawing_days):
    """根据出图制作时间调整节点排布（v3.7：拆分“开始出图”与“交付”）。

    将 v3.6 合并的“最终节点：出图与图纸归档”区间（end-N+1 ~ end）拆分为
    两个独立节点：“开始出图”落在 end-N+1（end - (N-1) 天），“交付”落在
    end，交付与整体周期最终日期重合；N=1 时两节点同日不判重合。原“图纸
    归档”节点不再单独保留（其内容并入“交付”节点说明）。
    drawing_days<=0 时不处理。
    """
    if not drawing_days or drawing_days <= 0:
        return node_map
    end_items = node_map.get(end, [])
    kept = []
    archive = None
    for item in end_items:
        if item[0] == "图纸归档" and archive is None:
            archive = item
        else:
            kept.append(item)
    node_map[end] = kept
    # v3.7：拆分为“开始出图”（end-N+1）与“交付”（end）
    span_start = end - timedelta(days=drawing_days) + timedelta(days=1)  # end-N+1
    out_content = ("按照院相关文件的要求完成出图（设计文件交付之前完成）；"
                   "开始出图日期为 {0}，交付日期为 {1}（交付与整体周期最终"
                   "日期重合，制作时间共 {2} 天，N=1 时两节点同日不判重合）。").format(
                       fmt_date_cn(span_start), fmt_date_cn(end), drawing_days)
    archive_content = archive[1] if archive is not None else ""
    node_map.setdefault(span_start, []).append(
        ("开始出图", out_content,
         "出图交付", "出图", None))
    delivery_content = (archive_content if archive_content
                        else "交付设计文件（与整体周期最终日期重合）。")
    node_map.setdefault(end, []).append(
        ("交付", delivery_content,
         "出图交付", "出图", None))
    return node_map


def adjust_review_meeting_schedule(node_map, start, end, drawing_days,
                                   avoid_holidays, holidays, workdays):
    """v3.2（v3.7 表述更新）：将“专业间核对、会审与会签”节点安排在“开始出图”之前。

    会审节点原按 None 百分比落在设计结束日；现改为 end-N（“开始出图”日期
    end-N+1 的前一天），保证会审在开始出图之前且两者不重叠。若 end-N 为
    休息日且启用了避开节假日，则向前取最近工作日（保证不晚于 end-N）；
    未启用避开节假日时按原休息日日期放置。drawing_days<=0（未启用出图
    制作）时会审保持落结束日。
    """
    if not drawing_days or drawing_days <= 0:
        return node_map
    review_title = "专业间核对、会审与会签"
    target = end - timedelta(days=drawing_days)
    if target < start:
        target = start
    # 先从节点图中摘除会审节点
    moved = []
    for d, items in list(node_map.items()):
        kept = []
        for item in items:
            if item[0] == review_title:
                moved.append(item)
            else:
                kept.append(item)
        if kept:
            node_map[d] = kept
        else:
            node_map.pop(d, None)
    if not moved:
        return node_map
    for item in moved:
        d = target
        if avoid_holidays and is_rest_day(d, holidays, workdays):
            cur = d - timedelta(days=1)
            while cur >= start and is_rest_day(cur, holidays, workdays):
                cur -= timedelta(days=1)
            if cur >= start and not is_rest_day(cur, holidays, workdays):
                d = cur  # 向前取最近工作日，不晚于 end-N
            else:
                d = target  # 范围内找不到工作日，保持 target
        node_map.setdefault(d, []).append(item)
    return node_map


def fmt_date_cn(d):
    """如 2026年01月05日，避免 strftime 中文编码问题。"""
    return "{0}年{1:02d}月{2:02d}日".format(d.year, d.month, d.day)


def fmt_date_short(d):
    return "{0:02d}-{1:02d}".format(d.month, d.day)


def estimate_row_height(cell_text, col_width):
    """v5.3.2(L3)：按文本实际宽度估算换行后的行高（每行约 14 磅）。

    中文/全角字符按 2 单位、英文/数字按 1 单位计算文本宽度，再按列宽折算
    换行行数，避免多语言译文或长文本被截断。
    """
    if not cell_text:
        return 20
    chars_per_line = max(int(col_width * 1.05), 2)  # 以英文字符为单位的单行容量（约等于列宽）
    lines = 0
    for seg in str(cell_text).split("\n"):
        width = 0
        for ch in seg:
            width += 2 if ord(ch) > 0x2E7F else 1  # 中文/全角按 2，英文/数字按 1
        lines += max(1, math.ceil(width / chars_per_line))
    return max(22, lines * 14 + 8)


def weekday_cn(d):
    """星期几中文名。"""
    return ("周一", "周二", "周三", "周四", "周五", "周六", "周日")[d.weekday()]


def collect_statistics(node_map, start, end, holidays, workdays, drawing_days):
    """统计项目信息：总天数/总周数/工作日/休息日/节点数/里程碑数/计划出图日期。

    返回 dict。总周数按日历周排布计算（含首尾不完整周）。
    """
    total_days = (end - start).days + 1
    work_day_count = 0
    rest_day_count = 0
    cur = start
    while cur <= end:
        if is_rest_day(cur, holidays, workdays):
            rest_day_count += 1
        else:
            work_day_count += 1
        cur += timedelta(days=1)
    # 日历周数：从开始日所在周一到结束日所在周日，跨过的周数
    week_start = start - timedelta(days=start.weekday())
    week_end = end + timedelta(days=6 - end.weekday())
    total_weeks = ((week_end - week_start).days + 1) // 7
    node_total = sum(
        1 for items in node_map.values()
        for it in items if it[4] != "_stage_end")  # v3.4：阶段结束标注不计入节点数
    milestone_total = sum(
        1 for items in node_map.values()
        for title, _, _, _, p in items if p != "_stage_end" and is_milestone(title))
    # 计划出图日期：v3.7 优先取“交付”节点日期（与整体周期最终日期重合）；
    # 无“交付”节点时退回含“出图”节点；再无则取结束日
    drawing_date = end
    delivery_found = False
    for d, items in node_map.items():
        for title, _, _, _, _ in items:
            if title == "交付":
                drawing_date = d
                delivery_found = True
                break
        if delivery_found:
            break
    if not delivery_found:
        for d, items in node_map.items():
            for title, _, _, _, _ in items:
                if "出图" in title:
                    drawing_date = d
                    break
            if drawing_date != end:
                break
    return {
        "total_days": total_days,
        "total_weeks": total_weeks,
        "work_days": work_day_count,
        "rest_days": rest_day_count,
        "node_total": node_total,
        "milestone_total": milestone_total,
        "drawing_date": drawing_date,
    }


def write_node_list_sheet(wb, node_map, start, end, fmt, stage_name, style="classic", languages=None, overlap_dates=None):
    """新增“节点清单”汇总工作表：按日期顺序列出全部节点。

    列：序号、日期、星期、所属阶段、节点名称、原定日期、迁移来源、重合状态、
    具体内容或章节编号、里程碑标记（v5.3.2：新增原定日期/迁移来源/重合状态三列）。
    fmt=full 时内容列显示具体配合内容；fmt=brief 时显示章节编号提示。
    style: v4.0 输出风格（classic/blue/green/warm），配色取自 STYLE_PALETTES。
    languages: v5.0 多语种代码列表；在标题区、表头、阶段名、节点名称后追加译文
    （内置词典覆盖，未覆盖词条保持中文不编造）。
    overlap_dates: v5.3.2 重合日期集合（普通重合 + 避开节假日调整重合），
    节点日期落入集合时“重合状态”列标注“重合”。
    """
    ws = wb.create_sheet("节点清单")

    _pals = load_style_palettes()
    pal = _pals.get(style, _pals["classic"])
    thin = Side(style="thin", color=pal["border"])
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor=pal["header"])
    header_font = Font(name="微软雅黑", size=11, bold=True, color=pal["header_text"])
    body_font = Font(name="微软雅黑", size=10, color="262626")
    milestone_font = Font(name="微软雅黑", size=10, bold=True, color=pal["accent"])
    title_font = Font(name="微软雅黑", size=16, bold=True, color=pal["title"])
    info_font = Font(name="微软雅黑", size=10, color="404040")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    wrap_left = Alignment(wrap_text=True, vertical="top", horizontal="left")
    alt_fill = PatternFill("solid", fgColor=pal["zebra"])
    milestone_fill = PatternFill("solid", fgColor="FFE699")

    # 标题区
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = "{0}　{1}".format(translate_text_cn(stage_name, languages),
                                translate_text_cn("节点清单汇总表", languages))
    c.font = title_font
    c.alignment = wrap_center
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:J2")
    c = ws["A2"]
    c.value = "{0}：{1}　{2}：{3} ~ {4}（共 {5} 天）　{6}：{7}".format(
        translate_text_cn("生成日期", languages), fmt_date_cn(date.today()),
        translate_text_cn("设计周期", languages), fmt_date_cn(start), fmt_date_cn(end),
        (end - start).days + 1, translate_text_cn("输出版本", languages),
        translate_text_cn("完整版" if fmt == "full" else "简洁版", languages))
    c.font = info_font
    c.alignment = wrap_center
    ws.row_dimensions[2].height = 20

    headers = [translate_text_cn(h, languages) for h in
               ["序号", "日期", "星期", "所属阶段", "节点名称", "原定日期", "迁移来源",
                "重合状态", "具体内容或章节编号", "里程碑标记"]]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = wrap_center
        cell.border = border_all
    ws.row_dimensions[3].height = 22

    widths = [6, 12, 8, 22, 26, 12, 10, 10, 70, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    seq = 1
    row = 4
    overlap_dates = overlap_dates or set()
    ordered_dates = sorted(node_map.keys())
    for d in ordered_dates:
        items = node_map[d]
        for title, content, stage, section, percent in items:
            milestone = is_milestone(title)
            # v5.3.2(L1)：解析迁移标注（原定日期/迁移来源）与重合状态
            orig_date = ""
            migrate_src = "—"
            _mk = "（由"
            _mi = content.find(_mk)
            if _mi >= 0:
                _rest = content[_mi + len(_mk):]
                _month, _, _tail = _rest.partition("月")
                _day = _tail.split("日", 1)[0] if "日" in _tail else ""
                if _month.isdigit() and _day.isdigit():
                    orig_date = "{0}月{1}日".format(_month, _day)
                    migrate_src = "节假日"
            overlap_mark = "重合" if d in overlap_dates else ""
            cell_no = ws.cell(row=row, column=1, value=seq)
            cell_date = ws.cell(row=row, column=2, value=fmt_date_cn(d))
            cell_week = ws.cell(row=row, column=3, value=weekday_cn(d))
            cell_stage = ws.cell(row=row, column=4,
                                 value=translate_text_cn(stage, languages))
            title_tr = translate_app_text(title, languages)
            if percent == "_stage_end":
                cell_name = ws.cell(row=row, column=5,
                                    value="【阶段结束】" + title_tr)
            else:
                cell_name = ws.cell(row=row, column=5,
                                    value=("【里程碑】" + title_tr) if milestone else title_tr)
            cell_orig = ws.cell(row=row, column=6, value=orig_date)
            cell_mig = ws.cell(row=row, column=7, value=migrate_src)
            cell_ov = ws.cell(row=row, column=8, value=overlap_mark)
            if fmt == "full":
                cell_content = ws.cell(row=row, column=9,
                                       value=translate_content_cn(content, languages))
            else:
                if section:
                    cell_content = ws.cell(row=row, column=9,
                                           value=translate_section_ref_cn(section, languages))
                else:
                    cell_content = ws.cell(row=row, column=9,
                                           value=translate_content_cn(content, languages))
            cell_mark = ws.cell(row=row, column=10,
                                value="【里程碑】" if milestone else "—")
            for cell in (cell_no, cell_date, cell_week, cell_stage, cell_name,
                         cell_orig, cell_mig, cell_ov, cell_content, cell_mark):
                cell.border = border_all
                cell.font = milestone_font if milestone and cell in (cell_name, cell_mark) else body_font
            for cell in (cell_no, cell_date, cell_week, cell_stage, cell_orig,
                         cell_mig, cell_ov, cell_mark):
                cell.alignment = wrap_center
            cell_name.alignment = wrap_left
            cell_content.alignment = wrap_left
            if milestone:
                for cell in (cell_no, cell_date, cell_week, cell_stage, cell_name,
                             cell_orig, cell_mig, cell_ov, cell_content, cell_mark):
                    cell.fill = milestone_fill
            elif seq % 2 == 0:
                for cell in (cell_no, cell_date, cell_week, cell_stage, cell_name,
                             cell_orig, cell_mig, cell_ov, cell_content, cell_mark):
                    cell.fill = alt_fill
            # 行高按内容估算
            est = estimate_row_height(str(cell_content.value), widths[8])
            ws.row_dimensions[row].height = max(20, est)
            seq += 1
            row += 1

    # 底部汇总行
    total = sum(
        1 for items in node_map.values()
        for it in items if it[4] != "_stage_end")  # v3.4：阶段结束标注不计入节点数
    stage_end_count = sum(
        1 for items in node_map.values()
        for it in items if it[4] == "_stage_end")
    milestone_count = sum(
        1 for items in node_map.values()
        for title, _, _, _, p in items if p != "_stage_end" and is_milestone(title))
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    extra = " 阶段结束标注 {0} 条（【阶段结束】行，不计入节点数）。".format(stage_end_count) if stage_end_count else ""
    c = ws.cell(row=row, column=1,
                value="合计：共 {0} 个节点，其中里程碑节点 {1} 个（已加【里程碑】标记并标黄高亮）。{2}".format(
                    total, milestone_count, extra))
    c.font = Font(name="微软雅黑", size=10, bold=True, color="1F3864")
    c.alignment = wrap_left
    ws.row_dimensions[row].height = 20

    # 打印设置
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:3"
    ws.freeze_panes = "A4"
    return ws


def write_work_list_sheets(wb, node_map, start, end, fmt, stage_name, style="classic", stage_key="preliminary", languages=None, extra_profs=None, overwrite_entries=None):
    """按专业写入专业作业清单 sheet（v4.1；v5.0 支持多语种译文）。

    表名即专业名（建筑/总平面/结构/给排水/暖通/电气/智能化/经济），格式参考
    「节点清单」：序号/日期/星期/所属阶段/章节编号/节点名称/内容/类型（发图/收图）。
    发图（提出）行淡橙 #FCE4D6，收图（接收）行淡蓝 #DDEBF7；同一节点中既有
    提出又有接收的内容已按段落拆分到对应专业行，类型列明确标注，不混色。
    数据从 node_map 按专业标签分组推导（见 _extract_work_entries），日期与
    主表一致（含休息日迁移/出图拆分结果），内容均为节点原文。
    v5.3.2(T2)：新增独立“章节编号”列，取节点所属章节编号（node_map 中 section），
    覆写节点等无章节来源的条目显示“—”。
    languages: v5.0 多语种代码列表；表名、表头、阶段名、节点名称、类型与
    内容按词典追加译文（未覆盖词条保持中文不编造）。
    返回实际生成的 sheet 数。
    """
    _pals = load_style_palettes()
    palette = _pals.get(style, _pals["classic"])
    entries = _extract_work_entries(node_map, stage_key, extra_profs, overwrite_entries)
    if not entries:
        return 0
    # v5.3.2(T2)：由 node_map 构建 (日期, 标题) → 章节编号 查找表
    section_lookup = {}
    for d, items in node_map.items():
        for title, content, stage, section, percent in items:
            if section:
                section_lookup.setdefault((d, title), section)
    header_fill = PatternFill("solid", fgColor=palette["header"])
    header_font = Font(name="微软雅黑", size=10, bold=True, color=palette["header_text"])
    title_font = Font(name="微软雅黑", size=13, bold=True, color=palette["title"])
    note_font = Font(name="微软雅黑", size=9, color="595959")
    send_fill = PatternFill("solid", fgColor=WORK_LIST_SEND_FILL)
    recv_fill = PatternFill("solid", fgColor=WORK_LIST_RECV_FILL)
    thin = Side(style="thin", color=palette.get("border", "808080"))
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap_left = Alignment(wrap_text=True, vertical="top", horizontal="left")
    center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    headers = [translate_text_cn(h, languages) for h in
               ["序号", "日期", "星期", "所属阶段", "章节编号", "节点名称", "内容/事项", "类型"]]
    base_profs = list(WORK_LIST_PROFESSIONS.get(stage_key, []))
    extra_profs = list(extra_profs or [])
    all_profs = base_profs + [p for p in extra_profs if p not in base_profs]
    made = 0
    for prof in all_profs:
        items = entries.get(prof)
        if not items:
            if prof in base_profs:
                continue
            # v5.0：自定义应用标准中出现的新专业，即使无具体提图/收图条目也生成工作表
            ws = wb.create_sheet(title=prof[:31])
            ws.merge_cells("A1:H1")
            c = ws["A1"]
            c.value = "{0}{1}（{2}）".format(
                translate_text_cn(prof, languages),
                translate_text_cn("专业作业清单", languages),
                translate_text_cn(stage_name, languages))
            c.font = title_font
            c.alignment = Alignment(vertical="center", horizontal="left")
            ws.row_dimensions[1].height = 26
            ws.merge_cells("A2:H2")
            c2 = ws["A2"]
            c2.value = ("说明：该专业出现在自定义应用标准/覆写节点模板的“涉及专业”列中，"
                        "但节点内容未包含该专业的提图/收图描述，暂无具体条目。")
            c2.font = note_font
            c2.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            ws.row_dimensions[2].height = 30
            for ci, h in enumerate(headers, 1):
                cell = ws.cell(row=3, column=ci, value=h)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = center
            ws.cell(row=4, column=1, value="——")
            for ci, w in enumerate([6, 12, 8, 22, 14, 26, 90, 8], 1):
                ws.column_dimensions[get_column_letter(ci)].width = w
            ws.freeze_panes = "A4"
            made += 1
            continue
        ws = wb.create_sheet(title=prof[:31])
        ws.merge_cells("A1:H1")
        c = ws["A1"]
        c.value = "{0}{1}（{2}）".format(
            translate_text_cn(prof, languages),
            translate_text_cn("专业作业清单", languages),
            translate_text_cn(stage_name, languages))
        c.font = title_font
        c.alignment = Alignment(vertical="center", horizontal="left")
        ws.row_dimensions[1].height = 26
        ws.merge_cells("A2:H2")
        c2 = ws["A2"]
        c2.value = ("说明：发图（提出）行底色淡橙（#{0}），收图（接收）行底色淡蓝（#{1}）；"
                    "内容摘自《建筑工程设计项目专业配合规定》节点原文，日期与主计划表一致。").format(
            WORK_LIST_SEND_FILL, WORK_LIST_RECV_FILL)
        c2.font = note_font
        c2.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        ws.row_dimensions[2].height = 30
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=ci, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center
        row = 4
        idx = 1
        for (direction, d, stage, title, content_text) in items:
            section_no = section_lookup.get((d, title), "")
            ws.cell(row=row, column=1, value=idx)
            ws.cell(row=row, column=2, value=fmt_date_cn(d))
            ws.cell(row=row, column=3, value="星期{0}".format("一二三四五六日"[d.weekday()]))
            ws.cell(row=row, column=4, value=translate_text_cn(stage, languages))
            ws.cell(row=row, column=5,
                    value=translate_section_ref_cn(section_no, languages) if section_no else "—")
            ws.cell(row=row, column=6, value=translate_app_text(title, languages))
            ws.cell(row=row, column=7, value=translate_content_cn(content_text, languages))
            ws.cell(row=row, column=8, value=translate_text_cn(direction, languages))
            fill = send_fill if direction == "发图" else recv_fill
            for ci in range(1, 9):
                cell = ws.cell(row=row, column=ci)
                cell.fill = fill
                cell.border = border
                cell.alignment = center if ci in (1, 2, 3, 5, 8) else wrap_left
            lines = 1 + content_text.count("\n")
            ws.row_dimensions[row].height = max(30, min(150, 15 * lines + 8))
            row += 1
            idx += 1
        for ci, w in enumerate([6, 12, 8, 22, 14, 26, 90, 8], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "A4"
        ws.print_title_rows = "1:3"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        made += 1
    return made



# =====================================================================
# v5.0：自定义应用标准 + 覆写节点（运行支持目录 / 模板导出 / 模板解析 / 覆写叠加）
# =====================================================================
DEFAULT_STANDARD_NAME = "中国建筑设计研究院专业配合规定"
RUNTIME_ROOT_NAME = "运行支持"
_RT_OVW = "覆写节点文件"
_STAGE_DIR_NAMES = {"preliminary": "初步设计", "construction": "施工图设计"}
TEMPLATE_HEADERS = ["标准名称", "涉及专业", "节点名称", "节点周期百分比（%）",
                    "开始于总周期的第几天（与百分比二选一）", "节点内容",
                    "发图专业", "收图专业", "章节编号", "里程碑（是/否）"]
OVERWRITE_HEADERS = ["节点名称", "节点周期百分比（%）",
                     "开始于总周期的第几天（与百分比二选一）", "节点内容",
                     "涉及专业", "发图专业", "收图专业"]


def _stage_dir_cn(stage_key):
    return _STAGE_DIR_NAMES.get(stage_key, "初步设计")


def _default_holidays_json():
    """v5.3.2：内置节假日/调休数据（2020-2030，含 2027-2030 预测值）拍平为
    holidays.json 的默认内容，格式 {"holidays": [...], "workdays": [...]}。
    clear_builtin=False 表示在程序内置数据基础上追加/对齐，不替换内置数据。"""
    return {
        "holidays": [d for year in sorted(_HOLIDAYS) for d in _HOLIDAYS[year]],
        "workdays": [d for year in sorted(_WORKDAYS) for d in _WORKDAYS[year]],
        "clear_builtin": False,
    }


def ensure_runtime_dirs(base_dir=None):
    """v5.0：首次运行建立运行支持文件夹（标准模板/覆写节点文件/应用标准文件），
    每类下按初步设计、施工图设计分目录。
    v5.3.2：运行支持目录下不存在 holidays.json 时自动生成默认节假日文件。
    返回根目录绝对路径。"""
    base = base_dir or program_dir()
    root = os.path.join(base, RUNTIME_ROOT_NAME)
    for kind in (_RT_OVW,):
        for sd in ("初步设计", "施工图设计"):
            os.makedirs(os.path.join(root, kind, sd), exist_ok=True)
    holidays_path = os.path.join(root, "holidays.json")
    if not os.path.exists(holidays_path):
        try:
            with open(holidays_path, "w", encoding="utf-8") as f:
                json.dump(_default_holidays_json(), f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return root


def _template_notes(kind_name, is_overwrite=False):
    notes = [
        "《{0}》填写说明（本说明区与表头请保留；数据行可增删、顺序不限）：".format(kind_name),
        "1. 标准名称：填写本模板对应的标准名称；同一文件建议统一，留空默认“{0}”。".format(DEFAULT_STANDARD_NAME),
        "2. 涉及专业：填写该节点涉及的专业，多个专业用顿号/逗号分隔；程序默认专业（建筑、总平面、结构、"
        "给排水、暖通、电气、智能化、经济）之外出现的新专业，生成表格时会自动建立对应 sheet 工作表并生成专业作业清单。",
        "3. 节点周期百分比（%）与“开始于总周期的第几天”：二选一填写。百分比填写数值（如 30 表示第 30%）；"
        "第几天填写数值（如 12 表示开始于总周期第 12 天，含首尾）。两者均留空时该节点落在周期结束日。",
        "4. 节点内容：填写该节点需要表达的具体配合内容（可留空）。",
        "5. 发图专业/收图专业：填写提图/收图专业（可留空）；专业作业清单将按节点内容自动推导。",
        "6. 章节编号：程序排布时按节点名称自动匹配，可留空。",
        "7. 里程碑（是/否）：仅“是”的节点按里程碑样式显示。",
        "8. 行序不限、数量不限：程序读取后按时间节点自动排序；本模板保存在运行支持\\{0}\\<阶段>\\ 下。".format(kind_name),
    ]
    return notes


def _write_template_workbook(wb, kind_name, headers, rows, col_widths):
    """v5.0：写模板工作表。第1行标题、第2-9行填写说明、第10行表头、第11行起数据。"""
    ws = wb.active
    ws.title = "模板"
    title_fill = PatternFill("solid", fgColor="1F3864")
    note_fill = PatternFill("solid", fgColor="FCE4D6")
    head_fill = PatternFill("solid", fgColor="DDEBF7")
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws.cell(row=1, column=1, value=kind_name)
    c.font = Font(name="微软雅黑", size=13, bold=True, color="FFFFFF")
    c.fill = title_fill
    c.alignment = Alignment(vertical="center", horizontal="left")
    ws.row_dimensions[1].height = 26
    notes = _template_notes(kind_name)
    for i, txt in enumerate(notes):
        r = 2 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(headers))
        cc = ws.cell(row=r, column=1, value=txt)
        cc.font = Font(name="微软雅黑", size=9)
        cc.fill = note_fill
        cc.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        ws.row_dimensions[r].height = 30 if i in (1, 2) else 18
    hdr_row = 2 + len(notes)
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=hdr_row, column=ci, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.fill = head_fill
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for ri, row in enumerate(rows, start=hdr_row + 1):
        for ci, v in enumerate(row, 1):
            ws.cell(row=ri, column=ci, value=v if v is not None else "")
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A{0}".format(hdr_row)


def export_overwrite_template(stage_key, base_dir=None, out_path=None):
    """v5.0：导出一份覆写节点模板 Excel（空模板 + 示例行）。返回保存路径。"""
    ensure_runtime_dirs(base_dir)
    stage_cn = _stage_dir_cn(stage_key)
    if out_path is None:
        out_dir = os.path.join(ensure_runtime_dirs(base_dir), _RT_OVW, stage_cn)
        out_path = os.path.join(out_dir, "覆写节点模板_{0}.xlsx".format(stage_cn))
    rows = [["示例：外审会议", "20", "", "示例内容，可删除", "", "", ""],
            ["示例：幕墙深化资料", "", "35", "示例内容，可删除", "幕墙", "幕墙", "建筑、结构"]]
    wb = Workbook()
    _write_template_workbook(wb, "覆写节点模板（{0}）".format(stage_cn), OVERWRITE_HEADERS,
                             rows, [26, 16, 16, 60, 16, 14, 18])
    wb.save(out_path)
    return out_path


def _locate_header_row(ws, key_headers):
    """在模板工作表中定位表头行（需同时包含若干表头关键词），返回列名->列号映射。"""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        vals = [(c.value if c.value is not None else "") for c in row]
        joined = [str(v) for v in vals]
        if all(any(k in str(v) for v in joined) for k in key_headers):
            return {str(v).strip(): idx + 1 for idx, v in enumerate(vals) if str(v).strip()}
    return None


def ensure_template_samples():
    """v4.5中国院版：首次运行自动生成覆写节点模板样例到运行支持目录（不存在时）。"""
    made = []
    try:
        ensure_runtime_dirs()
        for st in ("preliminary", "construction"):
            d = os.path.join(ensure_runtime_dirs(), _RT_OVW, _stage_dir_cn(st))
            os.makedirs(d, exist_ok=True)
            if not any(fn.lower().endswith(".xlsx") for fn in os.listdir(d)):
                export_overwrite_template(st)
                made.append(d)
    except Exception as e:
        print("警告：无法自动生成模板样例：{0}".format(e))
    return made


def parse_overwrite_xlsx(path, stage_key):
    """v5.0（v5.1 扩展）：读取覆写节点模板 xlsx。
    返回 dict：{name, nodes:[(percent_or_day, stage, title, content,
    send_profs, recv_profs), ...], extra_profs:[新专业有序列表], issues:[提示]}。
    send_profs/recv_profs 分别来自“发图专业”“收图专业”列（标准专业名列表，
    无该列或留空时为空列表）；内容保持原文（不进行多语种处理）。"""
    from openpyxl import load_workbook as _load
    wb = _load(path, data_only=True)
    ws = wb.active
    hdr_row = None
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        vals = [(c.value if c.value is not None else "") for c in row]
        joined = [str(v) for v in vals]
        if all(any(k in str(v) for v in joined) for k in ["节点名称", "节点周期百分比"]):
            hdr_row = row[0].row
            colmap = {str(v).strip(): idx + 1 for idx, v in enumerate(vals) if str(v).strip()}
            break
    if hdr_row is None:
        raise ValueError("无法识别覆写节点模板表头行")
    c_name = colmap.get("节点名称")
    c_percent = next((v for k, v in colmap.items() if "百分比" in k), None)
    c_day = next((v for k, v in colmap.items() if "第几天" in k), None)
    c_content = next((v for k, v in colmap.items() if "节点内容" in k), None)
    c_profs = next((v for k, v in colmap.items() if "涉及专业" in k), None)
    c_send = next((v for k, v in colmap.items() if "发图专业" in k), None)
    c_recv = next((v for k, v in colmap.items() if "收图专业" in k), None)
    stage_cn = _stage_dir_cn(stage_key)
    out = []
    profs_all = []
    issues = []
    seen = set()  # v5.3.3(U4)：重复节点检测
    default_profs = set(WORK_LIST_PROFESSIONS.get(stage_key, []))
    for row in ws.iter_rows(min_row=hdr_row + 1, max_row=ws.max_row, max_col=ws.max_column):
        vals = [c.value for c in row]
        title = vals[c_name - 1] if c_name and c_name <= len(vals) else None
        if title is None or str(title).strip() == "":
            continue
        title = str(title).strip()
        content = str(vals[c_content - 1]).strip() if c_content and c_content <= len(vals) and vals[c_content - 1] is not None else ""
        if not content:
            issues.append("节点《{0}》内容为空".format(title))
        if title in seen:
            issues.append("节点《{0}》重复出现，将叠加生效".format(title))
        seen.add(title)
        p = None
        if c_percent and c_percent <= len(vals) and vals[c_percent - 1] is not None and str(vals[c_percent - 1]).strip() != "":
            raw = str(vals[c_percent - 1]).strip().replace("%", "")
            try:
                p = float(raw)
            except ValueError:
                p = None
                issues.append("节点《{0}》的百分比无法解析（{1}），已忽略该百分比字段".format(title, raw))
            if isinstance(p, float) and (p < 0 or p > 100):
                issues.append("节点《{0}》的百分比 {1} 超出 0~100 范围，按实际值参与换算".format(title, p))
        elif c_day and c_day <= len(vals) and vals[c_day - 1] is not None and str(vals[c_day - 1]).strip() != "":
            try:
                p = ("day", int(float(str(vals[c_day - 1]).strip())))
            except ValueError:
                p = None
        if c_profs and c_profs <= len(vals) and vals[c_profs - 1] is not None:
            raw = str(vals[c_profs - 1]).strip()
            if raw:
                for tok in _split_prof_tokens(raw):
                    if tok and tok not in profs_all:
                        profs_all.append(tok)
        send_profs = []
        if c_send and c_send <= len(vals) and vals[c_send - 1] is not None:
            send_profs = _split_prof_tokens(str(vals[c_send - 1]).strip())
        recv_profs = []
        if c_recv and c_recv <= len(vals) and vals[c_recv - 1] is not None:
            recv_profs = _split_prof_tokens(str(vals[c_recv - 1]).strip())
        out.append((p, stage_cn, title, content, send_profs, recv_profs))
    if not out:
        raise ValueError("覆写节点模板中未读取到任何节点（请检查“节点名称”列是否有内容）")
    extra_profs = [p for p in profs_all if p not in default_profs]
    return {"name": os.path.basename(path), "nodes": out, "extra_profs": extra_profs, "issues": issues}


def _norm_template_nodes(node_list, period_days):
    """v5.0：把模板解析出的节点归一为数值百分比；返回 (归一列表, 紧密标志)。
    (“day”, n) 表示开始于总周期第 n 天，n>=周期天数视为节点过于紧密。
    v5.1：节点第 5、6 个元素（发图专业/收图专业列表，仅覆写节点模板有）原样透传，
    供专业作业清单按覆写节点内容填充收发图条目。"""
    out = []
    tight = False
    period = max(period_days - 1, 1)
    for node in node_list:
        p = node[0]
        if isinstance(p, tuple) and p and p[0] == "day":
            n = int(p[1])
            if n < 1:
                n = 1
            if n >= period_days:
                tight = True
            percent = (n - 1) / period * 100.0
        else:
            percent = p
            if percent is not None and percent >= 100:
                tight = True
        tail = tuple(node[4:])
        out.append((percent, node[1], node[2], node[3]) + tail)
    out.sort(key=lambda x: (x[0] is None, x[0] if x[0] is not None else 0))
    return out, tight


def add_overwrite_nodes(node_map, overwrite_nodes, start, end, drawing_days, round_mode="round"):
    """v5.0：把覆写节点叠加进 node_map。percent 标记 "_custom" 不参与普通重合判定、
    不因普通重合标浅红；与出图日期（end）重合时自动移至出图制作起始日前一天并返回提示。
    v5.1：节点元组可为 4 元组 (percent, stage, title, content) 或 6 元组
    (percent, stage, title, content, send_profs, recv_profs)。send_profs/recv_profs
    用于为专业作业清单生成收发图条目（发图行方向"发图"、收图行方向"收图"），
    内容与节点原文一致；覆写节点在“避开节假日”开启时也会被迁移（调用方保证
    migrate_rest_day_nodes 在本函数之后执行）。
    返回 (warnings, entries)；entries 为 {标准专业名: [(方向, 日期, 阶段, 标题, 内容), ...]}。
    """
    warnings = []
    entries = {}
    for node in overwrite_nodes:
        percent, stage, title, content = node[0], node[1], node[2], node[3]
        send_profs = node[4] if len(node) > 4 else []
        recv_profs = node[5] if len(node) > 5 else []
        d = percent_to_date(start, end, percent, round_mode)
        if d is None or d < start or d > end:
            continue
        if d == end and drawing_days and drawing_days > 0:
            d = end - timedelta(days=drawing_days)
            if d < start:
                d = start
            warnings.append("覆写节点《{0}》与出图日期重合，自动调整至出图前，请注意".format(title))
        node_map.setdefault(d, []).append((title, content, stage or "覆写节点", "", "_custom"))
        stage_lbl = stage or "覆写节点"
        for s in send_profs:
            if s:
                entries.setdefault(s, []).append(("发图", d, stage_lbl, title, content))
        for r in recv_profs:
            if r:
                entries.setdefault(r, []).append(("收图", d, stage_lbl, title, content))
    return warnings, entries


def translate_app_text(text, languages):
    """v5.0：多语种输出，未覆盖词条回退内置词典。"""
    if not languages:
        return text
    return translate_text_cn(text, languages)


def compute_plan_layout(stage_key, start, end, drawing_days=1,
                        avoid_holidays=False, overwrite_nodes=None,
                        round_mode="round"):
    """v5.3.3(U6/U8)：不落盘计算节点落地布局，返回 node_map（日期->items）。"""
    hf = None
    try:
        rt_dir = ensure_runtime_dirs()
        hp = os.path.join(rt_dir, "holidays.json")
        if os.path.exists(hp):
            hf = hp
    except Exception:
        hf = None
    holidays, workdays = load_holidays(hf)
    period_days = (end - start).days + 1
    if period_days <= 1:
        raise ValueError("结束日期必须晚于开始日期")
    nodes = PRELIMINARY_NODES if stage_key == "preliminary" else CONSTRUCTION_NODES
    ov_nodes = None
    if overwrite_nodes:
        ov_nodes, _t = _norm_template_nodes(overwrite_nodes, period_days)
    node_map = build_node_map(nodes, start, end, stage_key, round_mode)
    apply_stage_period_rules(node_map, start, end, stage_key, round_mode)
    apply_drawing_schedule(node_map, start, end, drawing_days)
    adjust_review_meeting_schedule(node_map, start, end, drawing_days,
                                   avoid_holidays, holidays, workdays)
    if ov_nodes:
        _w, _e = add_overwrite_nodes(node_map, ov_nodes, start, end, drawing_days, round_mode)
    if avoid_holidays:
        migrate_rest_day_nodes(node_map, start, end, holidays, workdays, skip_titles=("交付",))
    return node_map


def collect_preflight_issues(plan, cfg_preflight=None):
    """v5.3.3(U6)：生成前预检——返回 (重合列表, 紧密列表, 迁移列表)。"""
    pf = cfg_preflight or {}
    try:
        min_gap = max(0, int(pf.get("min_gap_days") or 0))
    except (TypeError, ValueError):
        min_gap = 0
    overlap_items = []
    for d, items in sorted(plan.items()):
        pcts = [it[4] for it in items if it[4] not in ("_custom", "_stage_end")]
        if len(set(pcts)) >= 2:
            overlap_items.append((d, [it[0] for it in items]))
    tight_items = []
    if min_gap > 0:
        dates = sorted(plan.keys())
        for i in range(1, len(dates)):
            gap = (dates[i] - dates[i - 1]).days
            if gap < min_gap:
                tight_items.append((dates[i - 1], dates[i], gap))
    migrate_items = []
    for d, items in sorted(plan.items()):
        for it in items:
            if it[1] and "迁移" in str(it[1]):
                migrate_items.append((d, it[0], it[1]))
    return overlap_items, tight_items, migrate_items


def generate_excel(stage_key, start, end, output_path, holidays, workdays,
                   project_name="", design_code="", avoid_holidays=False,
                   fmt="full", drawing_days=1, style="classic",
                   work_lists=False, languages=None,
                   overwrite_nodes=None, overwrite_extra_profs=None,
                   round_mode="round"):
    """生成 Excel 设计计划时间表。

    fmt: full=完整版（含各节点具体配合内容）；brief=简洁版（仅显示节点与章节编号）。
    drawing_days: 出图制作时间（天），>0 时将最终节点拆分为“开始出图”
                  （end-N+1）与“交付”（end）两个节点（v3.7，交付与整体周期
                  最终日期重合，N=1 同日不判重合），并把会审与会签节点安排在
                  “开始出图”之前（end-N，休息日且避开节假日时向前取工作日）；
                  <=0 时不处理。
    style: v4.0 输出风格（classic/blue/green/warm），配色取自 STYLE_PALETTES；
           休息日淡绿、重合浅红、斑马纹等硬性偏好各风格统一保留。
    work_lists: v4.1 是否生成「专业作业清单」sheet（默认 False）；True 时按
                所选设计阶段（stage_key）为每个专业新增一张工作表，发图行
                淡橙 #FCE4D6、收图行淡蓝 #DDEBF7，数据从节点数据按专业标签
                分组推导，日期与主表一致。
    languages: v5.0 多语种代码列表；在标题区、表头、阶段名、节点名称后追加译文
                （内置词典未覆盖的词条保持中文不编造；None/空列表=仅中文）。
    overwrite_nodes: v5.0 覆写节点列表 [(percent, stage, title, content[, send_profs, recv_profs]), ...]，
                叠加到标准/自定义应用标准之上；普通重合（不同百分比节点同日）不参与判定与标色，
                内容保持原文；与出图日期（end）重合时自动移至出图制作起始日前一天。
                v5.1：覆写节点在“避开节假日”开启时同样参与迁移（落休息日迁至最近工作日并注明原日期）；
                因避开节假日调整导致的同日重合会标浅红并提示“由于避开节假日调整导致节点重合，请注意”；
                模板“发图专业/收图专业”列用于为专业作业清单生成收发图条目（含新增专业）。
    overwrite_extra_profs: v5.0 覆写节点模板“涉及专业”列识别出的新专业列表，
                并入专业作业清单/工作表。
    v5.1 变更：
      * “避开节假日”同时作用于默认标准与覆写节点；
      * 因避开节假日调整导致节点同日重合时，对应日期格标浅红并注明
        “※ 由于避开节假日调整导致节点重合，请注意”；
      * 覆写模式下新增专业作业清单不再为空：按覆写节点“发图专业/收图专业”列
        填充收发图条目（内容与节点原文一致），且其它专业清单也体现新增专业收发图节点。
    返回 (文件路径, 节点数, 重合日期数, 统计dict, 提示列表 warnings)。
    """
    if fmt not in ("full", "brief"):
        fmt = "full"
    _pals = load_style_palettes()  # v5.3.3(U1)：内置+自定义样式合并表
    if style not in _pals:
        style = "classic"
    if stage_key == "preliminary":
        stage_name = "初步设计阶段"
        nodes = PRELIMINARY_NODES
        periods = PRELIMINARY_PERIODS
        period_symbol = "P"
    else:
        stage_name = "施工图设计阶段"
        nodes = CONSTRUCTION_NODES
        periods = CONSTRUCTION_PERIODS
        period_symbol = "W"

    # v5.0：设计周期含首尾（v3.7：end = start + (periodDays - 1)）
    period_days = (end - start).days + 1
    if period_days <= 1:
        raise ValueError("结束日期必须晚于开始日期")
    # v5.0：自定义应用标准/覆写节点归一化（day 模式转百分比）+ 周期紧密校验
    tight_period = False
    ov_nodes = None
    if overwrite_nodes:
        ov_nodes, t2 = _norm_template_nodes(overwrite_nodes, period_days)
        tight_period = tight_period or t2

    node_map = build_node_map(nodes, start, end, stage_key, round_mode)
    apply_stage_period_rules(node_map, start, end, stage_key, round_mode)  # v3.5 阶段周期表达规则（第四/五阶段）
    apply_drawing_schedule(node_map, start, end, drawing_days)
    adjust_review_meeting_schedule(node_map, start, end, drawing_days,
                                   avoid_holidays, holidays, workdays)
    # v5.1：覆写节点先叠加，再统一执行“避开节假日”迁移，
    # 使覆写节点同样按既定规则避开节假日。
    ow_warnings = []
    ov_entries = {}
    if ov_nodes:
        ow_warnings, ov_entries = add_overwrite_nodes(node_map, ov_nodes, start, end, drawing_days, round_mode)
    if avoid_holidays:
        migrate_rest_day_nodes(node_map, start, end, holidays, workdays, skip_titles=("交付",))
    warnings = list(ow_warnings)
    if tight_period:
        warnings.append("节点过于紧密，请注意计划安排合理性")

    # 重合检测（v3.2）：仅当同一日期存在两个及以上原始时间设定（百分比）不同的
    # 节点时才判定为异常重合；文档中同百分比多节点同日属正常设计；
    # 阶段结束标注（percent="_stage_end"）不参与判定。
    overlap_dates = set()
    for d, items in node_map.items():
        if len(items) < 2:
            continue
        pcts = [it[4] for it in items if it[4] not in ("_custom", "_stage_end")]
        if len(set(pcts)) >= 2:
            overlap_dates.add(d)
    overlap_count = len(overlap_dates)

    # v5.3.2：因“避开节假日”调整导致的节点同日重合（默认标准/覆写节点
    # 迁移后与迁至日原有节点重合）：同日至少存在一个带
    # “迁移而来”标注的迁移节点，且同时存在未迁移的原有节点，才判定为迁移
    # 导致重合；若当日节点全部从休息日整体迁移而来（迁至日原本无节点），
    # 不属于“迁移导致与迁至日节点重合”，不标浅红。
    holiday_overlap_dates = set()
    if avoid_holidays:
        for d, items in node_map.items():
            if len(items) < 2:
                continue
            moved = [it for it in items if it[1] and "迁移而来" in it[1]]
            non_moved = [it for it in items if not (it[1] and "迁移而来" in it[1])]
            if moved and non_moved:
                holiday_overlap_dates.add(d)
    holiday_overlap_count = len(holiday_overlap_dates)
    if holiday_overlap_count:
        warnings.append("检测到 {0} 个日期因避开节假日调整导致节点重合，请注意".format(holiday_overlap_count))

    # ---- 样式（v4.0：按合并样式表取色；休息日淡绿/重合浅红/斑马纹硬偏好保留）----
    pal = _pals.get(style, _pals["classic"])
    thin = Side(style="thin", color=pal["border"])
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor=pal["header"])
    header_font = Font(name="微软雅黑", size=11, bold=True, color=pal["header_text"])
    rest_fill = PatternFill("solid", fgColor="E2EFDA")   # 硬性偏好：淡绿色休息日/节假日
    work_fill = PatternFill("solid", fgColor="FFFFFF")   # 白色：工作日
    zebra_fill = PatternFill("solid", fgColor=pal["zebra"])  # 斑马纹行（工作日交替）
    overlap_fill = PatternFill("solid", fgColor="FFC7CE")  # 硬性偏好：浅红色节点重合
    title_font = Font(name="微软雅黑", size=16, bold=True, color=pal["title"])
    info_font = Font(name="微软雅黑", size=10, color="404040")
    date_font = Font(name="微软雅黑", size=10, bold=True, color=pal["title"])
    node_title_font = Font(name="微软雅黑", size=9, bold=True, color=pal["accent"])
    node_body_font = Font(name="微软雅黑", size=9, color="262626")
    milestone_cell_font = Font(name="微软雅黑", size=9, bold=True, color=pal["accent"])
    week_font = Font(name="微软雅黑", size=10, bold=True)
    note_font = Font(name="微软雅黑", size=9, color="595959")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    wrap_left = Alignment(wrap_text=True, vertical="top", horizontal="left")
    # v3.6：循环内逐格新建样式对象开销大（openpyxl 每次赋值都做去重/哈希），
    # 提取为复用常量，缩短“生成→询问”前的执行时间。
    week_col_fill = PatternFill("solid", fgColor=pal["week_col"])   # 周次列底色
    out_range_fill = PatternFill("solid", fgColor="F2F2F2")  # 范围外灰底
    node_cell_font = Font(name="微软雅黑", size=9, color=pal["title"])
    plain_cell_font = Font(name="微软雅黑", size=9, color="262626")

    wb = Workbook()
    ws = wb.active
    ws.title = "设计计划时间表"

    # 列宽：A 周次，B~H 周一~周日
    widths = [12, 20, 20, 20, 20, 20, 20, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- 标题区 ----
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = "{0}　{1}".format(translate_text_cn(stage_name, languages),
                                translate_text_cn("设计配合进度计划表", languages))
    c.font = title_font
    c.alignment = wrap_center
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:H2")
    c = ws["A2"]
    project_text = project_name.strip() if project_name else ""
    code_text = design_code.strip() if design_code else ""
    if not project_text:
        project_text = "——"
    if not code_text:
        code_text = "——"
    c.value = "{0}：{1}　　{2}：{3}".format(translate_text_cn("项目名称", languages),
                                           project_text,
                                           translate_text_cn("设计编号", languages),
                                           code_text)
    c.font = Font(name="微软雅黑", size=12, bold=True, color=pal["title"])
    c.alignment = wrap_center
    ws.row_dimensions[2].height = 22

    ws.merge_cells("A3:H3")
    c = ws["A3"]
    c.value = "{0}：{1}　　{2}：{3}　　{4}：{5}　　{6}：{7} 天".format(
        translate_text_cn("设计阶段", languages), translate_text_cn(stage_name, languages),
        translate_text_cn("开始日期", languages), fmt_date_cn(start),
        translate_text_cn("结束日期", languages), fmt_date_cn(end),
        translate_text_cn("设计周期", languages), period_days)
    c.font = info_font
    c.alignment = wrap_center
    ws.row_dimensions[3].height = 20

    ws.merge_cells("A4:H4")
    c = ws["A4"]
    avoid_text = ("{0}（{1}）".format(translate_text_cn("避开节假日", languages),
                                     translate_text_cn("节点迁移至最近工作日", languages))
                  if avoid_holidays else "不避开节假日")
    fmt_text = translate_text_cn("完整版" if fmt == "full" else "简洁版", languages)
    dd_text = ("{0} 天".format(drawing_days) if drawing_days and drawing_days > 0
               else translate_text_cn("未启用", languages))
    c.value = "{0}：{1}　　{2}：{3}　　{4}：{5}　　{6}：{7}　　（节点日期 = 开始日 + round(百分比 × 周期天数)）".format(
        translate_text_cn("编制日期", languages), fmt_date_cn(date.today()),
        translate_text_cn("节假日处理", languages), avoid_text,
        translate_text_cn("输出版本", languages), fmt_text,
        translate_text_cn("出图制作时间", languages), dd_text)
    c.font = info_font
    c.alignment = wrap_center
    ws.row_dimensions[4].height = 20

    # ---- 表头 ----
    header_row = 5
    headers = [translate_text_cn("周次", languages)] + \
        [translate_text_cn(h, languages) for h in ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = wrap_center
        cell.border = border_all
    ws.row_dimensions[header_row].height = 22

    # ---- 逐周排布 ----
    # 从开始日期所在周的周一起，到结束日期所在周的周日止
    week_start = start - timedelta(days=start.weekday())      # 周一
    week_end = end + timedelta(days=6 - end.weekday())        # 周日
    row = header_row + 1
    week_no = 1
    today = date.today()
    cur = week_start
    while cur <= week_end:
        # 周次列
        wc = ws.cell(row=row, column=1,
                     value="第{0}周\n{1} ~ {2}".format(week_no, fmt_date_short(cur),
                                                      fmt_date_short(cur + timedelta(days=6))))
        wc.font = week_font
        wc.alignment = wrap_center
        wc.border = border_all
        wc.fill = week_col_fill

        max_lines = 1
        for col in range(2, 9):  # 周一~周日
            d = cur + timedelta(days=col - 2)
            cell = ws.cell(row=row, column=col)
            cell.border = border_all
            if d < start or d > end:
                cell.fill = out_range_fill
                cell.alignment = wrap_center
                continue
            is_rest = is_rest_day(d, holidays, workdays)
            is_overlap = d in overlap_dates
            is_holiday_overlap = d in holiday_overlap_dates
            # 底色优先级：范围外灰 > 节点重合浅红（含避开节假日调整重合） > 休息日淡绿 > 工作日斑马纹（隔行交替）
            if is_overlap or is_holiday_overlap:
                cell.fill = overlap_fill
            elif is_rest:
                cell.fill = rest_fill
            else:
                cell.fill = zebra_fill if (week_no % 2 == 0) else work_fill
            lines = ["{0}  {1}".format(fmt_date_short(d),
                                       translate_text_cn("（休息日）", languages)
                                       if is_rest else translate_text_cn("（工作日）", languages))]
            if d == today:
                lines[0] += " ★"
            # v5.2：迁移重合不再输出整格提示，仅由迁移节点自身标注
            # “（由X月X日（节假日）迁移而来）”，未迁移节点不显示迁移来源标注。
            if is_overlap:
                lines.append("")
                lines.append(translate_text_cn("※ 由于整体周期过短，导致节点重合，需要人工协调", languages))
            # 节点内容
            cell_has_milestone = False
            if d in node_map:
                for title, content, stage, section, percent in node_map[d]:
                    lines.append("")
                    if percent == "_stage_end":
                        # v3.4：阶段周期结束标注（非节点，不加◆/里程碑）
                        lines.append("▷ " + translate_text_cn("阶段结束标注", languages) + "："
                                     + translate_text_cn(stage, languages) + "："
                                     + translate_app_text(title, languages))
                        if fmt != "brief":
                            lines.append(translate_text_cn(content, languages))
                    elif is_milestone(title):
                        cell_has_milestone = True
                        lines.append("◆ " + translate_text_cn("【里程碑】", languages) + " "
                                     + translate_text_cn(stage, languages) + "："
                                     + translate_app_text(title, languages))
                    else:
                        lines.append("◆ " + translate_text_cn(stage, languages) + "："
                                     + translate_app_text(title, languages))
                    if percent != "_stage_end" and fmt == "brief":
                        lines.append("内容详专业配合规定第{0}节{1}".format(
                            section.split(".")[0], section))
                    elif percent != "_stage_end":
                        lines.append(translate_text_cn(content, languages))
            text = "\n".join(lines)
            cell.value = text
            cell.alignment = wrap_left
            if cell_has_milestone:
                cell.font = milestone_cell_font
            elif d in node_map:
                cell.font = node_cell_font
            else:
                cell.font = plain_cell_font
            max_lines = max(max_lines, len(lines))
        ws.row_dimensions[row].height = max(28, max_lines * 13 + 8)
        cur += timedelta(days=7)
        row += 1
        week_no += 1

    # ---- 项目统计信息区（v3.1）----
    stats = collect_statistics(node_map, start, end, holidays, workdays, drawing_days)
    stat_row = row + 1
    ws.merge_cells(start_row=stat_row, start_column=1, end_row=stat_row, end_column=8)
    c = ws.cell(row=stat_row, column=1, value="【项目概况】")
    c.font = Font(name="微软雅黑", size=10, bold=True, color=pal["title"])
    c.alignment = wrap_left
    ws.row_dimensions[stat_row].height = 18

    stat_row += 1
    ws.merge_cells(start_row=stat_row, start_column=1, end_row=stat_row, end_column=8)
    c = ws.cell(row=stat_row, column=1,
                value="设计周期总天数：{0} 天　　总周数：{1} 周　　工作日：{2} 天　　休息日：{3} 天".format(
                    stats["total_days"], stats["total_weeks"],
                    stats["work_days"], stats["rest_days"]))
    c.font = Font(name="微软雅黑", size=10, bold=True, color="404040")
    c.alignment = wrap_left
    ws.row_dimensions[stat_row].height = 18

    stat_row += 1
    ws.merge_cells(start_row=stat_row, start_column=1, end_row=stat_row, end_column=8)
    c = ws.cell(row=stat_row, column=1,
                value="节点总数：{0} 个　　里程碑节点：{1} 个　　计划出图日期：{2}".format(
                    stats["node_total"], stats["milestone_total"],
                    fmt_date_cn(stats["drawing_date"])))
    c.font = Font(name="微软雅黑", size=10, bold=True, color="404040")
    c.alignment = wrap_left
    ws.row_dimensions[stat_row].height = 18

    # ---- 底部说明区 ----
    note_row = stat_row + 1

    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
    dd = drawing_days if drawing_days and drawing_days > 0 else 0
    c = ws.cell(row=note_row, column=1,
                value="【节点说明】节点日期由文档规定的“提出时间”百分比换算：{0} = {1} 天（设计周期），"
                      "日期 = 开始日期 + 四舍五入(百分比 × {0})。百分比非整天的由编制时协商确定。"
                      "周期输入方式：方式1 开始日期+结束日期；方式2 开始日期+设计计划周期（设计总时间，天），"
                      "end = start + (periodDays - 1) 天。"
                      "出图制作时间 {2} 天：最终节点拆分为“开始出图”（end-{2}+1）与“交付”（end）两个节点，"
                      "交付与整体周期最终日期重合（N=1 时两节点同日，不判重合）；"
                      "第五阶段/第七阶段的“专业间核对、会审与会签”安排在开始出图之前（end-{2}，"
                      "避开节假日时向前取最近工作日），两者不重叠；交付固定于最终日期，不参与避开节假日迁移。".format(
                          period_symbol, period_days, dd))
    c.font = note_font
    c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    ws.row_dimensions[note_row].height = 30

    note_row += 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
    c = ws.cell(row=note_row, column=1, value="【阶段时间周期】")
    c.font = Font(name="微软雅黑", size=9, bold=True)
    ws.row_dimensions[note_row].height = 18

    note_row += 1
    period_text = "；".join("{0}：{1}".format(n, t) for n, t in periods)
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
    c = ws.cell(row=note_row, column=1, value=period_text)
    c.font = note_font
    c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    ws.row_dimensions[note_row].height = max(18, (len(period_text) // 120 + 1) * 16)

    note_row += 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
    c = ws.cell(row=note_row, column=1,
                value="【阶段周期表达规则（v3.5）】凡阶段标注百分比起止周期（如 76%W～82%W、xx%P～yy%P）"
                      "且阶段内具体内容无独立精确时间节点的，阶段内容全部在周期开始日期格表达，"
                      "周期结束日期格标注“X阶段结束”（▷ 阶段结束标注）。适用阶段：施工图第四阶段"
                      "（管道综合与对图，71%W～76%W）、第五阶段（提清资料，76%W～82%W）；初步设计各阶段"
                      "节点均有独立百分比或为非百分比标注，不适用。该标注参与“避开节假日”迁移："
                      "落休息日时迁移至最近工作日并注明原日期；不参与节点重合判定（所在日期格与其它"
                      "节点同格时不标浅红、不提示“由于整体周期过短，导致节点重合，需要人工协调”），"
                      "不计入节点总数与里程碑统计。")
    c.font = note_font
    c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    ws.row_dimensions[note_row].height = 34

    note_row += 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
    if avoid_holidays:
        c = ws.cell(row=note_row, column=1,
                    value="【图例】淡绿色底 = 法定休息日 / 节假日（含调休补班处理）；浅红色底 = 两个及以上时间设定（百分比）"
                          "不同的节点重合（提示：由于整体周期过短，导致节点重合，需要人工协调；文档中同百分比多节点同日"
                          "属正常设计，不标浅红），或因避开节假日调整导致节点重合（浅红，仅迁移节点标注"
                          "“由X月X日（节假日）迁移而来”，未迁移节点不标注）；"
                          "白色底 = 工作日；浅蓝色底（隔行交替） = 斑马纹，便于阅读；★ = 编制当日；"
                          "加粗深红色节点文本带【里程碑】前缀 = 里程碑节点（专业协调会/会审会签/图纸归档/出图/交付）；"
                          "▷ 阶段结束标注 = 阶段周期结束标注（如“第四阶段（管道综合与对图）结束”），不参与"
                          "重合判定；"
                          "节点内容中“由X月X日（节假日）迁移而来”表示该节点原定休息日、已迁移至最近工作日"
                          "（避开节假日时默认标准/自定义应用标准/覆写节点/自定义节点均参与迁移）。")
    else:
        c = ws.cell(row=note_row, column=1,
                    value="【图例】淡绿色底 = 法定休息日 / 节假日（含调休补班处理）；浅红色底 = 两个及以上时间设定（百分比）"
                          "不同的节点重合（提示：由于整体周期过短，导致节点重合，需要人工协调；文档中同百分比多节点同日"
                          "属正常设计，不标浅红）；白色底 = 工作日；浅蓝色底（隔行交替） = 斑马纹，便于阅读；★ = 编制当日；"
                          "加粗深红色节点文本带【里程碑】前缀 = 里程碑节点（专业协调会/会审会签/图纸归档/出图/交付）；"
                          "▷ 阶段结束标注 = 阶段周期结束标注（如“第四阶段（管道综合与对图）结束”），不参与"
                          "重合判定。")
    c.font = note_font
    ws.row_dimensions[note_row].height = 32

    note_row += 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
    c = ws.cell(row=note_row, column=1,
                value="【编制依据】本表依据{0}、{1}编制。节点时间、配合内容均摘自该文档第二节/第三节原文；"
                      "“P”为合同确定的初步设计周期、“W”为合同确定的施工图设计周期，计量单位均为天。".format(DOC_TITLE, DOC_SECTION[stage_key]))
    c.font = Font(name="微软雅黑", size=9, bold=True, color="7F3F00")
    c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    ws.row_dimensions[note_row].height = 40

    if ov_nodes:
        note_row += 1
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
        c = ws.cell(row=note_row, column=1,
                    value="【覆写节点】已叠加 {0} 个覆写节点（内容保持原文，不进行多语种翻译；"
                          "与标准节点或其他覆写节点普通重合时不提示、不标色；"
                          "启用“避开节假日”时同样参与迁移，因迁移调整导致的重合会标浅红，"
                          "仅迁移节点标注“由X月X日（节假日）迁移而来”。".format(len(ov_nodes)))
        c.font = Font(name="微软雅黑", size=9, bold=True, color="7F3F00")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        ws.row_dimensions[note_row].height = 18
    if tight_period:
        note_row += 1
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
        c = ws.cell(row=note_row, column=1,
                    value="【周期校验】节点过于紧密，请注意计划安排合理性。")
        c.font = Font(name="微软雅黑", size=9, bold=True, color="C00000")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        ws.row_dimensions[note_row].height = 18

    note_row += 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
    lang_summary = ""
    if languages:
        lang_names = "、".join(LANG_OPTIONS[c] for c in languages)
        lang_summary = "　多语种译文：{0}（词典未覆盖词条保持中文）".format(lang_names)
    c = ws.cell(row=note_row, column=1,
                value="【软件信息】设计计划时间表生成器 v4.5中国院版　软件作者：JaJa+AI　输出风格：{0}（{1}）{2}{3}".format(
                    style, _pals.get(style, _pals["classic"])["desc"],
                    "　专业作业清单：包含" if work_lists else "", lang_summary))
    c.font = Font(name="微软雅黑", size=9, bold=True, color=pal["title"])
    c.alignment = wrap_center
    ws.row_dimensions[note_row].height = 18

    # ---- 节点清单汇总表（v3.1；v4.0 传入风格；v5.3.2 传入重合集合）----
    overlap_union = overlap_dates | holiday_overlap_dates
    write_node_list_sheet(wb, node_map, start, end, fmt, stage_name, style, languages,
                          overlap_dates=overlap_union)

    # ---- 专业作业清单表（v4.1：勾选/指定 --work-lists 时按专业新增 sheet）----
    if work_lists:
        extra_profs = []
        for p in (overwrite_extra_profs or []):
            if p not in extra_profs:
                extra_profs.append(p)
        write_work_list_sheets(wb, node_map, start, end, fmt, stage_name,
                               style, stage_key, languages, extra_profs=extra_profs,
                               overwrite_entries=ov_entries)

    # ---- 打印与冻结设置（主表）----
    ws.freeze_panes = "B6"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:5"
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    wb.save(output_path)
    # v5.1：保存后立即释放工作簿句柄，避免文件被占用导致随后自动打开失败
    try:
        wb.close()
    except Exception:
        pass
    return output_path, len(node_map), overlap_count, stats, warnings

# =====================================================================
# 五、命令行 / 交互入口
# =====================================================================

def choose_output_path(default_name):
    """弹出系统保存对话框（tkinter）选择保存位置。

    GUI 不可用（无 tkinter / 无显示环境）时回退为提示用户手动输入保存路径。
    返回 (路径, 是否使用了系统对话框)；用户取消保存时返回 (None, True)。
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.update()
        try:
            path = filedialog.asksaveasfilename(
                title="选择 Excel 保存位置",
                initialfile=default_name,
                defaultextension=".xlsx",
                filetypes=[("Excel 工作簿", "*.xlsx")],
            )
        finally:
            root.destroy()
        if path:
            return path, True
        return None, True
    except Exception:
        p = input("无法弹出系统保存窗口，请输入保存路径（含文件名，留空使用默认“{0}”）：".format(default_name)).strip()
        if not p:
            p = os.path.join(os.getcwd(), default_name)
        return p, False


def run_gui(args, cfg):
    """v5.0 GUI 图形交互界面（tkinter，默认启动路径）。

    界面简洁、设计师风格：窗口顶部展示软件名称/作者 JaJa+AI/版本 v5.0；
    所有参数均在界面选择或输入（项目名称、设计编号、设计阶段、周期输入方式、
    开始/结束日期或总周期天数、出图制作时间、避开节假日、简洁版开关、表格
    风格、多语种选项、输出保存位置）；点击「生成时间表」后若未指定保存位置
    则弹出保存对话框，最终返回参数字典；用户取消/关闭窗口返回 None。
    GUI 模式严禁混用 input()，确认与提醒统一用 tkinter messagebox。
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    # ---- 从 config.json 记忆预填 ----
    cfg_style = cfg.get("style", "classic")
    _pals = load_style_palettes()  # v5.3.3(U1)：内置+自定义样式合并表
    if not isinstance(cfg_style, str) or cfg_style not in _pals:
        cfg_style = "classic"
    cfg_fmt = cfg.get("format", "full")
    if cfg_fmt not in ("full", "brief", "both"):
        cfg_fmt = "full"
    try:
        cfg_dd = int(cfg.get("drawing_days", 1))
    except (TypeError, ValueError):
        cfg_dd = 1
    # v5.3.2(A4/U4)：取整模式、阶段与日期记忆预填
    cfg_round = cfg.get("round_mode", "round")
    if cfg_round not in ("floor", "ceil", "round"):
        cfg_round = "round"
    cfg_stage = cfg.get("stage_key", "preliminary")
    if cfg_stage not in ("preliminary", "construction"):
        cfg_stage = "preliminary"
    cfg_start_date = cfg.get("start_date") or ""
    cfg_end_date = cfg.get("end_date") or ""

    result = {}

    def on_generate():
        project_name = var_project.get().strip()
        design_code = var_code.get().strip()
        stage_key = stage_var.get()
        if stage_key not in ("preliminary", "construction"):
            messagebox.showwarning("输入有误", "请选择设计阶段（初步设计 / 施工图设计）。")
            return
        start = parse_date_arg(ent_start.get().strip())
        if start is None:
            messagebox.showwarning("输入有误", "开始日期格式错误，应为 YYYY-MM-DD。")
            return
        end = None
        period_days = None
        if mode_var.get() == "2":
            try:
                period_days = int(ent_period.get().strip())
            except ValueError:
                messagebox.showwarning("输入有误", "设计总周期天数必须为整数。")
                return
            if period_days <= 0:
                messagebox.showwarning("输入有误", "设计总周期天数必须为正整数。")
                return
            end = start + timedelta(days=period_days - 1)
        else:
            end = parse_date_arg(ent_end.get().strip())
            if end is None:
                messagebox.showwarning("输入有误", "结束日期格式错误，应为 YYYY-MM-DD。")
                return
        if end <= start:
            messagebox.showwarning("输入有误", "结束日期必须晚于开始日期。")
            return
        try:
            drawing_days = int(ent_drawing.get().strip() or "1")
        except ValueError:
            messagebox.showwarning("输入有误", "出图制作时间必须为整数。")
            return
        if drawing_days < 0:
            messagebox.showwarning("输入有误", "出图制作时间不能为负数。")
            return
        # v5.3.2(A2)：出图制作时间上限校验（拒绝并提示合理范围）
        dd_err = validate_drawing_days(drawing_days, start, end)
        if dd_err:
            messagebox.showwarning("输入有误", dd_err)
            return
        fmt = fmt_var.get()
        if fmt not in ("full", "brief", "both"):
            fmt = "full"
        style = style_var.get()
        if style not in _pals:
            style = "classic"
        output_path = var_output.get().strip()

        if (end - start).days + 1 < 15:
            if not messagebox.askyesno("周期过短",
                                       "设计周期仅 {0} 天，节点可能大量重合。\n是否继续？".format((end - start).days + 1)):
                return
        # v5.3.2(L2)：周期过长（>730 天）确认弹窗
        if (end - start).days + 1 > 730:
            if not messagebox.askyesno("周期过长",
                                       "设计周期长达 {0} 天（超过两年），节点分布将非常稀疏。\n是否继续？".format(
                                           (end - start).days + 1), parent=root):
                return
        # v5.3.2(A3)：项目起止日期跨越内置数据未收录年份时弹窗警告
        ymin, ymax = builtin_holiday_year_range()
        if start.year < ymin or end.year > ymax:
            messagebox.showwarning(
                "节假日数据预警",
                "项目起止日期（{0} ~ {1}）超出内置节假日数据收录范围（{2}~{3} 年），\n"
                "相关年份节假日/调休数据可能不完整，请留意计划日期合理性\n"
                "（可通过运行支持目录 holidays.json 补充）。".format(
                    start.isoformat(), end.isoformat(), ymin, ymax),
                parent=root)

        if not output_path:
            stage_tag = "初步设计" if stage_key == "preliminary" else "施工图设计"
            suffix = "" if fmt == "full" else ("_简洁版" if fmt == "brief" else "_完整版")
            default_name = default_name_for(stage_tag, start, end, suffix)
            output_path = filedialog.asksaveasfilename(
                title="选择 Excel 保存位置",
                initialfile=default_name,
                defaultextension=".xlsx",
                filetypes=[("Excel 工作簿", "*.xlsx")],
                parent=root)
            if not output_path:
                messagebox.showinfo("已取消", "已取消保存，未生成文件。")
                return

        # v5.3.3(U6)：生成前重合/紧密/迁移预检（豁免项与阈值取 config.preflight）
        if not _run_preflight():
            return
        result.update({
            "stage_key": stage_key, "start": start, "end": end,
            "period_days": period_days,
            "project_name": project_name, "design_code": design_code,
            "avoid_holidays": bool(avoid_var.get()),
            "fmt": fmt, "drawing_days": drawing_days, "style": style,
            "work_lists": bool(worklists_var.get()),
            "languages": normalize_languages(
                [code for code, v in var_langs.items() if v.get()]),
            "output_path": output_path,
            "overwrite_nodes": [dict(ow) for ow in loaded_overwrites],
            "round_mode": round_var.get(),
        })
        root.destroy()

    def on_mode_change():
        if mode_var.get() == "2":
            ent_end.configure(state="disabled", bg="#F0F0F0")
            ent_period.configure(state="normal", bg="white")
        else:
            ent_end.configure(state="normal", bg="white")
            ent_period.configure(state="disabled", bg="#F0F0F0")

    def on_style_change(*_):
        s = style_var.get()
        if s in _pals:
            lbl_style_desc.configure(text=_pals[s]["desc"])

    def browse_output():
        d = filedialog.asksaveasfilename(
            title="选择 Excel 保存位置",
            initialfile="设计计划时间表.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            parent=root)
        if d:
            var_output.set(d)

    # ---- 窗口骨架（简洁设计师风格：深蓝横幅 + 浅灰分组卡片）----
    root = tk.Tk()
    root.title("设计计划时间表生成器 v4.5中国院版")
    root.resizable(True, True)  # v5.3.2(U6)：允许缩放窗口
    # v5.3.2(U6 修复2)：不再写死固定初始尺寸（原 900x720），
    # 初始窗口大小改由下方内容自适应逻辑决定（打开即全览全部信息与按钮）；
    # 仅保留最小尺寸，避免滚动 Canvas 内容区塌缩为 1x1。
    root.minsize(760, 580)
    root.configure(bg="#F4F6F8")
    try:
        root.attributes("-topmost", True)
        root.lift()
        root.update_idletasks()
        root.attributes("-topmost", False)
    except Exception:
        pass

    f_banner = tk.Frame(root, bg="#1F3864")
    f_banner.pack(fill="x")
    tk.Label(f_banner, text="设计计划时间表生成器", bg="#1F3864", fg="white",
             font=("微软雅黑", 16, "bold")).pack(pady=(14, 0))
    tk.Label(f_banner, text="作者：JaJa+AI　｜　版本：v4.5中国院版", bg="#1F3864", fg="#D9E2F3",
             font=("微软雅黑", 9)).pack(pady=(2, 12))

    # v5.3.2(U6)：内容区域支持滚动条，低分辨率/放大后控件仍可见
    # v5.3.2(U6 修复)：Canvas 显式指定初始 width/height，防止请求尺寸 1x1 导致内容区不可见
    _canvas = tk.Canvas(root, bg="#F4F6F8", highlightthickness=0, width=880, height=660)
    _vbar = tk.Scrollbar(root, orient="vertical", command=_canvas.yview)
    _canvas.configure(yscrollcommand=_vbar.set)
    _vbar.pack(side="right", fill="y")
    _canvas.pack(side="left", fill="both", expand=True)

    f_main = tk.Frame(_canvas, bg="#F4F6F8")
    _win_id = _canvas.create_window((0, 0), window=f_main, anchor="nw")

    def _on_canvas_configure(_e):
        _canvas.configure(scrollregion=_canvas.bbox("all"))
        # v5.3.2(U6 修复)：宽度取 Canvas 实际宽度，避免误用 f_main 事件宽度
        # （初始 1px）导致内容帧被锁死为 1px、内容不可见
        _canvas.itemconfigure(_win_id, width=_canvas.winfo_width())

    # v5.3.2(U6 修复)：同时监听 f_main 与 canvas 的 <Configure>，
    # 确保 canvas 初始布局/窗口缩放时 f_main 宽度同步为 canvas 宽度
    f_main.bind("<Configure>", _on_canvas_configure)
    _canvas.bind("<Configure>", _on_canvas_configure)
    _canvas.bind_all("<MouseWheel>", lambda e: _canvas.yview_scroll(int(-e.delta / 120), "units"))

    def _card(title):
        fr = tk.LabelFrame(f_main, text=title, font=("微软雅黑", 10, "bold"),
                           bg="#F4F6F8", fg="#1F3864", padx=10, pady=6)
        return fr

    # 项目信息
    fr_proj = _card("项目信息")
    fr_proj.pack(fill="x", pady=(0, 6))
    tk.Label(fr_proj, text="项目名称：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=0, column=0, sticky="e", pady=3)
    var_project = tk.StringVar(value=(cfg.get("project_name") or ""))
    ent_project = tk.Entry(fr_proj, textvariable=var_project, width=46, font=("微软雅黑", 9))
    ent_project.grid(row=0, column=1, sticky="w", pady=3)
    tk.Label(fr_proj, text="设计编号：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=1, column=0, sticky="e", pady=3)
    var_code = tk.StringVar(value=(cfg.get("design_code") or ""))
    ent_code = tk.Entry(fr_proj, textvariable=var_code, width=46, font=("微软雅黑", 9))
    ent_code.grid(row=1, column=1, sticky="w", pady=3)

    # 设计周期
    fr_cyc = _card("设计周期")
    fr_cyc.pack(fill="x", pady=(0, 6))
    tk.Label(fr_cyc, text="设计阶段：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=0, column=0, sticky="e", pady=3)
    stage_var = tk.StringVar(value=cfg_stage)  # v5.3.2(U4)：预填上次使用的设计阶段
    tk.Radiobutton(fr_cyc, text="初步设计", variable=stage_var, value="preliminary",
                   bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=0, column=1, sticky="w", pady=3)
    tk.Radiobutton(fr_cyc, text="施工图设计", variable=stage_var, value="construction",
                   bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=0, column=2, sticky="w", pady=3)
    tk.Label(fr_cyc, text="周期输入方式：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=1, column=0, sticky="e", pady=3)
    mode_var = tk.StringVar(value="1")
    tk.Radiobutton(fr_cyc, text="方式1：开始 + 结束日期", variable=mode_var, value="1",
                   bg="#F4F6F8", font=("微软雅黑", 9), command=on_mode_change).grid(row=1, column=1, columnspan=2, sticky="w", pady=3)
    tk.Radiobutton(fr_cyc, text="方式2：开始 + 设计总周期天数", variable=mode_var, value="2",
                   bg="#F4F6F8", font=("微软雅黑", 9), command=on_mode_change).grid(row=2, column=1, columnspan=2, sticky="w", pady=3)
    tk.Label(fr_cyc, text="开始日期：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=3, column=0, sticky="e", pady=3)
    ent_start = tk.Entry(fr_cyc, width=16, font=("微软雅黑", 9))
    ent_start.insert(0, cfg_start_date)  # v5.3.2(U4)：预填上次使用的开始日期
    ent_start.grid(row=3, column=1, sticky="w", pady=3)
    tk.Label(fr_cyc, text="YYYY-MM-DD", bg="#F4F6F8", fg="#888888", font=("微软雅黑", 8)).grid(row=3, column=2, sticky="w", pady=3)
    tk.Label(fr_cyc, text="结束日期：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=4, column=0, sticky="e", pady=3)
    ent_end = tk.Entry(fr_cyc, width=16, font=("微软雅黑", 9))
    ent_end.insert(0, cfg_end_date)  # v5.3.2(U4)：预填上次使用的结束日期
    ent_end.grid(row=4, column=1, sticky="w", pady=3)
    tk.Label(fr_cyc, text="YYYY-MM-DD", bg="#F4F6F8", fg="#888888", font=("微软雅黑", 8)).grid(row=4, column=2, sticky="w", pady=3)
    tk.Label(fr_cyc, text="总周期天数：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=5, column=0, sticky="e", pady=3)
    ent_period = tk.Entry(fr_cyc, width=16, font=("微软雅黑", 9), state="disabled", bg="#F0F0F0")
    ent_period.grid(row=5, column=1, sticky="w", pady=3)
    tk.Label(fr_cyc, text="天（含首尾）", bg="#F4F6F8", fg="#888888", font=("微软雅黑", 8)).grid(row=5, column=2, sticky="w", pady=3)

    # 生成选项
    fr_opt = _card("生成选项")
    fr_opt.pack(fill="x", pady=(0, 6))
    tk.Label(fr_opt, text="出图制作时间：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=0, column=0, sticky="e", pady=3)
    ent_drawing = tk.Entry(fr_opt, width=8, font=("微软雅黑", 9))
    ent_drawing.insert(0, str(cfg_dd if cfg_dd >= 0 else 1))
    ent_drawing.grid(row=0, column=1, sticky="w", pady=3)
    tk.Label(fr_opt, text="天（默认 1；0 表示不启用）", bg="#F4F6F8", fg="#888888", font=("微软雅黑", 8)).grid(row=0, column=2, sticky="w", pady=3)
    avoid_var = tk.BooleanVar(value=bool(cfg.get("avoid_holidays", False)))
    tk.Checkbutton(fr_opt, text="避开节假日（节点迁移至最近工作日）", variable=avoid_var,
                   bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=1, column=1, columnspan=2, sticky="w", pady=3)
    # v5.3.2(U3)：输出版本三选一（完整版 / 简洁版 / 完整+简洁）
    tk.Label(fr_opt, text="输出版本：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=2, column=0, sticky="e", pady=3)
    fmt_var = tk.StringVar(value=cfg_fmt)
    f_fmt = tk.Frame(fr_opt, bg="#F4F6F8")
    f_fmt.grid(row=2, column=1, columnspan=2, sticky="w", pady=3)
    tk.Radiobutton(f_fmt, text="完整版", variable=fmt_var, value="full",
                   bg="#F4F6F8", font=("微软雅黑", 9)).pack(side="left")
    tk.Radiobutton(f_fmt, text="简洁版", variable=fmt_var, value="brief",
                   bg="#F4F6F8", font=("微软雅黑", 9)).pack(side="left", padx=(8, 0))
    tk.Radiobutton(f_fmt, text="完整+简洁", variable=fmt_var, value="both",
                   bg="#F4F6F8", font=("微软雅黑", 9)).pack(side="left", padx=(8, 0))
    # v5.3.2(A4)：取整模式下拉（默认 round，保持现有行为）
    tk.Label(fr_opt, text="取整模式：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=3, column=0, sticky="e", pady=3)
    round_var = tk.StringVar(value=cfg_round)
    round_menu = tk.OptionMenu(fr_opt, round_var, *("floor", "ceil", "round"))
    round_menu.configure(font=("微软雅黑", 9), bg="white")
    round_menu.grid(row=3, column=1, sticky="w", pady=3)
    tk.Label(fr_opt, text="日期百分比换算取整方式（floor 向下 / ceil 向上 / round 四舍五入）",
             bg="#F4F6F8", fg="#888888", font=("微软雅黑", 8)).grid(row=3, column=2, sticky="w", pady=3)
    worklists_var = tk.BooleanVar(value=bool(cfg.get("work_lists", False)))
    tk.Checkbutton(fr_opt, text="是否包含专业作业清单（勾选后按专业新增 sheet：发图淡橙、收图淡蓝）", variable=worklists_var,
                   bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=4, column=1, columnspan=2, sticky="w", pady=3)
    tk.Label(fr_opt, text="表格风格：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=5, column=0, sticky="e", pady=3)
    style_var = tk.StringVar(value=cfg_style)
    style_var.trace_add("write", on_style_change)
    style_menu = tk.OptionMenu(fr_opt, style_var, *list(_pals.keys()))
    style_menu.configure(font=("微软雅黑", 9), bg="white")
    style_menu.grid(row=5, column=1, sticky="w", pady=3)
    lbl_style_desc = tk.Label(fr_opt, text=_pals[cfg_style]["desc"], bg="#F4F6F8",
                              fg="#555555", font=("微软雅黑", 8), wraplength=330, justify="left")
    lbl_style_desc.grid(row=5, column=2, sticky="w", pady=3)
    tk.Label(fr_opt, text="保存位置：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=6, column=0, sticky="e", pady=3)
    var_output = tk.StringVar(value="")
    ent_output = tk.Entry(fr_opt, textvariable=var_output, width=38, font=("微软雅黑", 9))
    ent_output.grid(row=6, column=1, sticky="w", pady=3)
    tk.Button(fr_opt, text="浏览…", command=browse_output, font=("微软雅黑", 9)).grid(row=6, column=2, sticky="w", pady=3)

    # 覆写节点（v5.0）：导出模板 / 加载模板 / 恢复初始状态
    loaded_overwrites = []   # [{"path","name","nodes"}]

    def refresh_lang_state():
        # v5.0：多语种输出，语言选项全部可选
        for code, v in var_langs.items():
            cb_langs[code].configure(state="normal")

    def _refresh_custom_panel():
        for w in fr_ovw_inner.winfo_children():
            w.destroy()
        if not loaded_overwrites:
            tk.Label(fr_ovw_inner, text="（未加载覆写节点模板）", bg="#F4F6F8",
                     fg="#888888", font=("微软雅黑", 8)).pack(anchor="w", pady=1)
        for ow in loaded_overwrites:
            row = tk.Frame(fr_ovw_inner, bg="#F4F6F8")
            row.pack(anchor="w", pady=1)
            tk.Label(row, text="{0}（{1} 个节点）".format(ow["name"], len(ow["nodes"])),
                     bg="#F4F6F8", font=("微软雅黑", 8)).pack(side="left")
            tk.Button(row, text="×", font=("微软雅黑", 8, "bold"), width=2,
                      command=lambda t=ow: _remove_overwrite(t)).pack(side="left", padx=(6, 0))

    def _remove_overwrite(ow):
        if ow in loaded_overwrites:
            loaded_overwrites.remove(ow)
        _refresh_custom_panel()
        refresh_lang_state()

    def _show_parse_report(issues, label):
        # v5.3.3(U4)：解析报告弹窗
        if not issues:
            messagebox.showinfo("解析报告", "{0}：解析正常，无异常项。".format(label), parent=root)
            return
        msg = "{0} 解析发现 {1} 项异常：\n\n".format(label, len(issues))
        for it in issues[:80]:
            msg += "· " + it + "\n"
        if len(issues) > 80:
            msg += "\n……共 {0} 项".format(len(issues))
        messagebox.showwarning("解析报告", msg, parent=root)

    def _preview_plan_now():
        # v5.3.3(U8)：即时刷新预览（返回 node_map，不落盘）
        stage_key = stage_var.get()
        if stage_key not in ("preliminary", "construction"):
            return None
        start = parse_date_arg(ent_start.get().strip())
        if start is None:
            return None
        if mode_var.get() == "2":
            try:
                period_days = int(ent_period.get().strip())
            except ValueError:
                return None
            if period_days <= 0:
                return None
            end = start + timedelta(days=period_days - 1)
        else:
            end = parse_date_arg(ent_end.get().strip())
            if end is None:
                return None
        if end <= start:
            return None
        try:
            drawing_days = int(ent_drawing.get().strip() or "1")
        except ValueError:
            drawing_days = 1
        # v5.3.6(B16 同源修复)：覆写节点从“文件描述 dict 列表”展平为节点元组列表，
        # 与 CLI gui_opts 处理保持一致；此前直接把 dict 列表传入 compute_plan_layout，
        # 在 _norm_template_nodes 中 node[0] 触发 KeyError(0)，导致生成前预检
        # 弹窗“无法执行生成前预检: 0”且表格无法生成。
        ov_flat = []
        for _ow in loaded_overwrites:
            _src = _ow.get("name") or os.path.basename(_ow.get("path") or "")
            for _nd in _ow.get("nodes") or []:
                ov_flat.append((_nd[0], _nd[1], _src + "：" + _nd[2]) + tuple(_nd[3:]))
        return compute_plan_layout(
            stage_key, start, end, drawing_days=drawing_days,
            avoid_holidays=bool(avoid_var.get()),
            overwrite_nodes=ov_flat,
            round_mode=round_var.get())

    def _run_preflight():
        # v5.3.3(U6)：生成前重合/紧密/迁移预检（豁免项与阈值取 config.preflight）
        try:
            plan = _preview_plan_now()
        except Exception as e:
            messagebox.showwarning("预检失败", "无法执行生成前预检：{0}".format(e), parent=root)
            return False
        if not plan:
            return True
        pf = cfg.get("preflight") or {}
        skip_custom = bool(pf.get("skip_custom", True))
        overlap_items, tight_items, migrate_items = collect_preflight_issues(plan, pf)
        if skip_custom:
            overlap_items = [(d, titles) for d, titles in overlap_items
                             if not all(str(t).startswith("[自定义]") for t in titles)]
        lines = []
        if overlap_items:
            lines.append("发现 {0} 组节点重合（同一天多个节点）：".format(len(overlap_items)))
            for d, titles in overlap_items[:20]:
                lines.append("  {0}  {1}".format(d.isoformat(), " / ".join(str(t) for t in titles[:6])))
            if len(overlap_items) > 20:
                lines.append("  ……共 {0} 组".format(len(overlap_items)))
        if tight_items:
            lines.append("发现 {0} 处紧密节点（间隔 < {1} 天）：".format(
                len(tight_items), pf.get("min_gap_days", 2)))
            for d1, d2, gap in tight_items[:20]:
                lines.append("  {0} → {1}（{2} 天）".format(d1.isoformat(), d2.isoformat(), gap))
            if len(tight_items) > 20:
                lines.append("  ……共 {0} 处".format(len(tight_items)))
        if migrate_items:
            lines.append("发现 {0} 个迁移节点（休息日迁移）：".format(len(migrate_items)))
            for d, t, src in migrate_items[:20]:
                lines.append("  {0}  {1}（{2}）".format(d.isoformat(), t, src))
            if len(migrate_items) > 20:
                lines.append("  ……共 {0} 个".format(len(migrate_items)))
        if not lines:
            return True
        msg = "\n".join(lines) + "\n\n是否继续生成？（可在 config.json 的 preflight 中调整 min_gap_days / skip_custom）"
        return messagebox.askyesno("生成前预检", msg, parent=root)

    def _load_overwrite(stage_key_menu):
        dlg_dir = os.path.join(ensure_runtime_dirs(), _RT_OVW, _stage_dir_cn(stage_key_menu))
        fps = filedialog.askopenfilenames(title="选择覆写节点模板（可多选，复选叠加）",
                                          initialdir=dlg_dir,
                                          filetypes=[("Excel 工作簿", "*.xlsx")], parent=root)
        if not fps:
            return
        for fp in fps:
            try:
                parsed = parse_overwrite_xlsx(fp, stage_key_menu)
            except Exception as e:
                messagebox.showerror("加载失败", "{0}\n{1}".format(os.path.basename(fp), e), parent=root)
                continue
            loaded_overwrites.append({"path": fp, "name": os.path.basename(fp),
                                      "nodes": parsed["nodes"], "extra_profs": parsed["extra_profs"]})
            _show_parse_report(parsed["issues"], "覆写节点文件“{0}”".format(os.path.basename(fp)))
        _refresh_custom_panel()
        refresh_lang_state()
        if loaded_overwrites:
            ov_profs = []
            for ow in loaded_overwrites:
                for p in (ow.get("extra_profs") or []):
                    if p not in ov_profs:
                        ov_profs.append(p)
            prof_msg = ("\n识别到新专业：{0}，将自动生成对应工作表与作业清单。".format("、".join(ov_profs))
                        if ov_profs else "")
            messagebox.showinfo("已加载覆写节点", "已叠加 {0} 个覆写节点模板（内容保持原文不翻译；"
                                "与出图日期重合自动调整至出图前）{1}".format(
                                    len(loaded_overwrites), prof_msg), parent=root)

    def _restore_initial():
        if not messagebox.askyesno("恢复初始状态",
                                   "将清除已加载的覆写节点，恢复内置默认标准"
                                   "（中国建筑设计研究院专业配合规定）。\n是否继续？", parent=root):
            return
        loaded_overwrites.clear()
        _refresh_custom_panel()
        refresh_lang_state()
        messagebox.showinfo("已恢复", "已恢复初始状态：内置默认标准，无覆写节点。", parent=root)

    def _export_tmpl(stage_key_menu):
        try:
            fp = export_overwrite_template(stage_key_menu)
            messagebox.showinfo("已导出", "覆写节点模板已保存：\n{0}".format(fp), parent=root)
        except Exception as e:
            messagebox.showerror("导出失败", str(e), parent=root)

    def open_custom_menu():
        # v4.5中国院版：覆写节点管理向导（导出模板 / 加载模板 / 恢复初始状态）
        menu = tk.Toplevel(root)
        menu.title("覆写节点管理（v4.5中国院版）")
        menu.resizable(False, False)
        menu.configure(bg="#F4F6F8")
        menu.transient(root)
        menu.grab_set()
        tk.Label(menu, text="覆写节点管理（v4.5中国院版）", bg="#1F3864", fg="white",
                 font=("微软雅黑", 11, "bold")).pack(fill="x", pady=(0, 6))
        f_body = tk.Frame(menu, bg="#F4F6F8")
        f_body.pack(fill="both", expand=True, padx=12, pady=6)
        tk.Label(f_body, text="请选择设计阶段：", bg="#F4F6F8", font=("微软雅黑", 9)).pack(anchor="w", pady=2)
        menu_stage = tk.StringVar(value="preliminary")
        f_stage = tk.Frame(f_body, bg="#F4F6F8")
        f_stage.pack(anchor="w")
        tk.Radiobutton(f_stage, text="初步设计", variable=menu_stage, value="preliminary",
                       bg="#F4F6F8", font=("微软雅黑", 9)).pack(side="left")
        tk.Radiobutton(f_stage, text="施工图设计", variable=menu_stage, value="construction",
                       bg="#F4F6F8", font=("微软雅黑", 9)).pack(side="left", padx=(12, 0))

        def _mk_btn(text, desc, cmd):
            tk.Button(f_body, text=text, command=cmd, font=("微软雅黑", 9),
                      width=42, anchor="w", padx=6).pack(anchor="w", pady=(2, 0))
            tk.Label(f_body, text=desc, bg="#F4F6F8", fg="#888888",
                     font=("微软雅黑", 8), wraplength=470, justify="left").pack(anchor="w", pady=(0, 4))

        tk.Label(f_body, text="── 覆写节点操作 ──", bg="#F4F6F8", fg="#7F3F00",
                 font=("微软雅黑", 9, "bold")).pack(anchor="w", pady=(6, 0))
        _mk_btn("1. 导出覆写设置模板",
                "导出一份“覆写节点模板”空表（含示例行），保存到 运行支持\覆写节点文件\<阶段>\；"
                "用户编辑后作为覆写节点叠加。", lambda: _export_tmpl(menu_stage.get()))
        _mk_btn("2. 加载覆写节点模板",
                "选择 运行支持\覆写节点文件\<阶段>\ 模板加载（可多选叠加）；加载后弹出解析报告，"
                "含非法百分比/越界/重复/空内容明细。", lambda: _load_overwrite(menu_stage.get()))
        _mk_btn("3. 恢复初始状态",
                "清除已加载的覆写节点，恢复内置默认标准（中国建筑设计研究院专业配合规定）。",
                _restore_initial)
        tk.Button(f_body, text="关闭", command=menu.destroy, font=("微软雅黑", 9)).pack(pady=(8, 4))

    fr_custom = _card("覆写节点")
    fr_custom.pack(fill="x", pady=(0, 6))
    tk.Button(fr_custom, text="覆写节点管理…", command=open_custom_menu, bg="#7F3F00", fg="white",
              font=("微软雅黑", 9, "bold"), padx=10).grid(row=0, column=0, sticky="w", pady=3)
    tk.Label(fr_custom, text="导出覆写模板/加载覆写节点、恢复初始状态",
             bg="#F4F6F8", fg="#555555", font=("微软雅黑", 8)).grid(row=0, column=1, sticky="w", padx=(8, 0), pady=3)
    tk.Label(fr_custom, text="已加载覆写节点：", bg="#F4F6F8", font=("微软雅黑", 9)).grid(row=1, column=0, sticky="ne", pady=3)
    fr_ovw_inner = tk.Frame(fr_custom, bg="#F4F6F8")
    fr_ovw_inner.grid(row=1, column=1, sticky="w", pady=3)
    _refresh_custom_panel()

    # 多语种选项（v5.0）
    fr_lang = _card("多语种选项")
    fr_lang.pack(fill="x", pady=(0, 6))
    tk.Label(fr_lang, text="勾选需要输出译文的语言（可多选；表格在中文基础上追加所选语言译文，"
                           "内置词典未覆盖的词条保持中文不编造）：",
             bg="#F4F6F8", font=("微软雅黑", 8), fg="#555555", wraplength=580, justify="left").grid(
        row=0, column=0, columnspan=2, sticky="w", pady=3)
    cfg_langs = normalize_languages(cfg.get("languages", []))
    var_langs = {}
    cb_langs = {}
    lang_codes = list(LANG_OPTIONS.keys())
    for i, code in enumerate(lang_codes):
        var_langs[code] = tk.BooleanVar(value=(code in cfg_langs))
        cb = tk.Checkbutton(fr_lang, text="{0}（{1}）".format(LANG_OPTIONS[code], code),
                            variable=var_langs[code], bg="#F4F6F8", font=("微软雅黑", 9))
        cb.grid(row=1 + i // 2, column=i % 2, sticky="w", pady=2)
        cb_langs[code] = cb

    # 底部按钮
    f_btn = tk.Frame(root, bg="#F4F6F8")
    f_btn.pack(side="bottom", fill="x", padx=12, pady=(4, 12))
    # 修复 U6 布局顺序：f_btn 须先于 canvas 参与 pack 分配，否则被 canvas(expand) 挤压成 1px 不可见。
    # 将 canvas 重新 pack 到最后，保持“banner 顶部 + f_btn 底部 + canvas 填充剩余”的布局。
    _canvas.pack_forget()
    _canvas.pack(side="left", fill="both", expand=True)
    tk.Button(f_btn, text="生成时间表", command=on_generate, bg="#1F3864", fg="white",
              font=("微软雅黑", 11, "bold"), padx=18, pady=4).pack(side="left", expand=True)
    tk.Button(f_btn, text="退出", command=root.destroy, font=("微软雅黑", 9),
              padx=12, pady=4).pack(side="right")

    root.protocol("WM_DELETE_WINDOW", root.destroy)
    on_mode_change()
    on_style_change()
    refresh_lang_state()
    # ---- v5.3.2(U6 修复2)：初始窗口尺寸按实际内容自适应，打开即全览全部信息与底部按钮 ----
    # 公式：总高 = 装饰高(标题栏+边框) + banner 请求高 + 内容高(canvas bbox 兜底 f_main 请求高)
    #        + 底部按钮区高 + 底部余量；总宽 = canvas 宽 + 左右边距 + 滚动条宽。
    # 屏幕放不下时取屏幕可用尺寸（任务栏/边距兜底），滚动条仍作为兜底保留。
    root.update_idletasks()
    try:
        _deco_h = root.winfo_rooty() - root.winfo_y()          # 标题栏等装饰高度
        _banner_h = f_banner.winfo_reqheight()                 # 深蓝横幅请求高度
        _bbox = _canvas.bbox("all")
        _content_h = (_bbox[3] if _bbox and len(_bbox) >= 4 else 0) or 0
        if _content_h <= 0:
            _content_h = f_main.winfo_reqheight()              # 兜底：内容帧请求高度
        _btn_h = f_btn.winfo_height() or f_btn.winfo_reqheight()  # 底部按钮区高度
        _margin = 10                                           # 底部余量(8-12px)
        _want_h = _deco_h + _banner_h + _content_h + _btn_h + _margin
        _scr_h = root.winfo_screenheight() - 50                # 屏幕可用高度兜底（任务栏/边距）
        _final_h = _want_h if _want_h <= _scr_h else _scr_h
        _canvas_w = _canvas.winfo_width() or 880
        _want_w = _canvas_w + 24 + 17                          # canvas + 左右边距(12*2) + 滚动条
        _scr_w = root.winfo_screenwidth() - 50                 # 屏幕可用宽度兜底
        _final_w = _want_w if _want_w <= _scr_w else _scr_w
        root.geometry("{0}x{1}".format(int(_final_w), int(_final_h)))
        root.update_idletasks()
    except Exception:
        pass

    # 自动化验证钩子：仅当环境变量 GUI_AUTOTEST=1 时，进入事件循环后进程内断言
    # 初始状态内容全览（scrollregion 高 <= canvas 可视高、底部按钮在窗口可视区内），
    # 输出实测数字后自动退出；正常使用不触发。
    if os.environ.get("GUI_AUTOTEST"):
        def _autotest_check():
            import json as _json
            try:
                root.update_idletasks()
                root.update()          # 强制完成窗口映射与布局，避免未映射导致高度误测
                _sr = _canvas.cget("scrollregion") or ""
                _parts = _sr.split()
                _sr_h = int(_parts[3]) if len(_parts) == 4 else 0
                _vis_h = _canvas.winfo_height()
                _win_w = root.winfo_width()
                _win_h = root.winfo_height()
                _root_bottom = root.winfo_rooty() + _win_h
                _btn_rows = []
                for _b in f_btn.winfo_children():
                    _txt = _b.cget("text")
                    _top = _b.winfo_rooty()
                    _bh = _b.winfo_height()
                    _btn_rows.append([_txt, _top, _bh, _top + _bh <= _root_bottom])
                _out = {
                    "window": [_win_w, _win_h],
                    "f_banner_h": f_banner.winfo_height(),
                    "f_btn_h": f_btn.winfo_height(),
                    "canvas_vis_h": _vis_h,
                    "scrollregion_h": _sr_h,
                    "scroll_ok": _sr_h <= _vis_h,
                    "content_h_used": _content_h,
                    "buttons": _btn_rows,
                }
                print("AUTOTEST_METRICS " + _json.dumps(_out, ensure_ascii=False))
                sys.stdout.flush()
                # 模拟缩小窗口：滚动条应重新生效（scrollregion 高 > canvas 可视高）
                root.geometry("760x580")
                root.update_idletasks()
                root.update()
                _sr2 = _canvas.cget("scrollregion") or ""
                _p2 = _sr2.split()
                _sr2_h = int(_p2[3]) if len(_p2) == 4 else 0
                _vis2 = _canvas.winfo_height()
                print("AUTOTEST_SHRINK " + _json.dumps({
                    "canvas_vis_h": _vis2,
                    "scrollregion_h": _sr2_h,
                    "scroll_usable": _sr2_h > _vis2,
                }, ensure_ascii=False))
                sys.stdout.flush()
            except Exception as _e:
                print("AUTOTEST_ERROR " + repr(_e))
                sys.stdout.flush()
            root.destroy()
        root.after(2000, _autotest_check)

    root.mainloop()
    return result if result else None


def parse_args(argv):
    p = argparse.ArgumentParser(
        description="依据《建筑工程设计项目专业配合规定》(2015版) 生成可编辑的 Excel 设计计划时间表（v4.5中国院版，作者 JaJa+AI；默认启动 GUI 图形界面）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  python generate_design_schedule.py\n"
               "  python generate_design_schedule.py --cli\n"
               "  python generate_design_schedule.py --version\n"
               "  python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-02-28 --style blue\n"
               "  python generate_design_schedule.py --stage 1 --start 2026-01-05 --period-days 55 --style green --work-lists\n"
               "  python generate_design_schedule.py --stage 施工图 --start 2026-03-01 --end 2026-05-31 --output D:\\plan.xlsx\n"
               "  python generate_design_schedule.py --project 测试项目 --code 2026-01 --stage 1 --start 2026-02-10 --end 2026-02-28 --avoid-holidays yes\n"
               "  python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-01-18 --format brief --drawing-days 1\n"
               "  python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-01-18 --format both --output D:\\plan.xlsx\n"
               "  python generate_design_schedule.py --stage 1 --start 2026-01-05 --end 2026-02-28 --languages en,ja,ko（v5.0：输出英语/日语/韩语译文）")
    p.add_argument("--project", help="项目名称（可选；交互模式也会询问，可沿用 config.json 记忆）")
    p.add_argument("--code", help="设计编号（可选；交互模式也会询问，可沿用 config.json 记忆）")
    p.add_argument("--avoid-holidays", nargs="?", const="yes",
                   choices=["yes", "no", "y", "n", "1", "0", "true", "false"],
                   help="是否避开节假日：把落在休息日的节点迁移到最近工作日并注明原日期。取值 yes/no/y/n/1/0/true/false；"
                        "单独使用 --avoid-holidays 等同于 yes；不传时交互模式会询问，默认沿用 config.json 记忆")
    p.add_argument("--stage", help="设计阶段：1 / 初步 / preliminary 表示初步设计阶段；2 / 施工图 / construction 表示施工图设计阶段")
    p.add_argument("--start", help="设计阶段开始日期，格式 YYYY-MM-DD")
    p.add_argument("--end", help="设计阶段结束日期，格式 YYYY-MM-DD（方式1：与 --start 搭配）")
    p.add_argument("--period-days", type=int, default=None,
                   help="设计计划周期（设计总时间，天），格式：正整数（方式2：与 --start 搭配，"
                        "自动计算 end = start + (periodDays - 1) 天，周期含首尾；传 0 表示不启用）")
    p.add_argument("--style", metavar="STYLE",
                   help="表格风格名称（内置 classic/blue/green/warm，或 config.json 中自定义样式名）；"
                        "不传时默认沿用 config.json 记忆，无记忆则 classic")
    wl = p.add_mutually_exclusive_group()
    wl.add_argument("--work-lists", dest="work_lists", action="store_true",
                    help="v4.1：包含专业作业清单（勾选后按专业新增 sheet：初步设计 8 专业、施工图 7 专业，"
                         "发图（提出）行淡橙 #FCE4D6、收图（接收）行淡蓝 #DDEBF7；缺省不含，可沿用 config.json 记忆）")
    wl.add_argument("--no-work-lists", dest="work_lists", action="store_false",
                    help="v4.1：不包含专业作业清单（缺省行为；与 --work-lists 互斥，优先于 config.json 记忆）")
    p.set_defaults(work_lists=None)
    p.add_argument("--languages", default=None,
                   help="v5.0：多语种选项，逗号分隔的语言代码（如 en,ja,ko；可用代码：en=英语、"
                        "ja=日语、ko=韩语、fr=法语、de=德语、es=西班牙语、pt=葡萄牙语、ru=俄语、"
                        "ar=阿拉伯语、it=意大利语、vi=越南语、th=泰语）。生成的 Excel 在中文基础上"
                        "追加所选语言译文（内置词典未覆盖的词条保持中文不编造）。不传时默认沿用 "
                        "config.json 记忆，无记忆则仅中文")
    p.add_argument("--cli", action="store_true",
                   help="强制命令行交互模式（v4.0）：默认启动 GUI 图形界面；--cli 或环境变量 DSS_NO_GUI=1 时跳过 GUI，"
                        "回退为命令行交互")
    p.add_argument("--version", action="store_true", help="打印版本号并退出（v5.0）")
    p.add_argument("--output", help="输出 xlsx 文件路径（默认自动生成到当前目录或弹出保存对话框）")
    p.add_argument("--holidays-file", help="自定义节假日 JSON 文件路径（可选）")
    p.add_argument("--format", choices=["full", "brief", "both"],
                   help="输出版本：full=完整版（含各节点具体配合内容）；brief=简洁版（仅显示节点与章节编号，"
                        "如“内容详专业配合规定第2节2.2.1”）；both=同时生成完整版与简洁版两个 xlsx。不传时交互模式会询问，"
                        "默认沿用 config.json 记忆")
    p.add_argument("--drawing-days", type=int, default=None,
                   help="出图制作时间（天），默认沿用 config.json 记忆（默认 1）；v3.7 起将最终节点拆分为“开始出图”"
                        "（end-N+1）与“交付”（end）两个节点，交付与设计阶段截止日期一致，N=1 时两节点同日不判重合。"
                        "传 0 表示不启用")
    p.add_argument("--round-mode", dest="round_mode", choices=["floor", "ceil", "round"],
                   default=None,
                   help="v5.3.2(A4)：日期百分比换算取整方式（floor=向下取整 / ceil=向上取整 / "
                        "round=四舍五入，默认 round）。不传时优先沿用 config.json 记忆，缺省 round")
    p.add_argument("--overwrite-file", action="append", metavar="XLSX",
                   help="v5.0：加载覆写节点模板 xlsx（可多次传入实现复选叠加；内容保持原文不翻译；"
                        "与出图日期重合自动调整至出图前）")
    p.add_argument("--open", nargs="?", const="yes",
                   choices=["yes", "no", "y", "n", "1", "0", "true", "false"],
                   help="生成成功后是否立即打开生成的 Excel（Windows 用 os.startfile，Linux 用 xdg-open）。"
                        "取值 yes/no；单独使用 --open 等同于 yes；不传时交互模式会询问")
    p.add_argument("--no-config", action="store_true",
                   help="不读取也不写入 config.json（用于测试/隔离配置记忆）")
    return p.parse_args(argv)


def parse_stage(s):
    s = str(s).strip().lower()
    if s in ("1", "初步", "初步设计", "初步设计阶段", "preliminary"):
        return "preliminary"
    if s in ("2", "施工图", "施工图设计", "施工图设计阶段", "construction"):
        return "construction"
    return None


def parse_format(s):
    s = str(s).strip().lower()
    if s in ("1", "full", "完整", "完整版", "f"):
        return "full"
    if s in ("2", "brief", "简洁", "简洁版", "b"):
        return "brief"
    if s in ("3", "both", "同时", "同时生成", "同时生成完整版与简洁版"):
        return "both"
    return None


def parse_date_arg(s):
    try:
        return date.fromisoformat(str(s).strip())
    except ValueError:
        return None


def parse_yes_no(s, default=False):
    s = str(s).strip().lower()
    if s in ("y", "yes", "1", "true"):
        return True
    if s in ("n", "no", "0", "false"):
        return False
    return default


def validate_drawing_days(drawing_days, start, end):
    """v5.3.2(A2)：出图制作时间上限校验。

    当 drawing_days >= 项目周期天数（含首尾）时，出图节点将排布到计划范围之外，
    静默生成错误表格。返回错误消息字符串；合法时返回 None。
    """
    if drawing_days is None or drawing_days <= 0:
        return None
    period = (end - start).days + 1
    if drawing_days >= period:
        return ("出图制作时间（{0} 天）不得大于等于项目周期（{1} 天，含首尾），"
                "否则“开始出图”节点将超出计划范围；请调整为 1~{2} 天（或 0 表示不启用）。"
                .format(drawing_days, period, period - 1))
    return None


def holiday_year_coverage(start, end, holidays, workdays):
    """v5.3.3(U5)：返回周期内各年份休息日覆盖情况与年份缺口列表。"""
    years = range(start.year, end.year + 1)
    rows = []
    for y in years:
        h_cnt = sum(1 for d in holidays if d.year == y)
        w_cnt = sum(1 for d in workdays if d.year == y)
        builtin = bool(_HOLIDAYS.get(y))
        rows.append((y, h_cnt, w_cnt, builtin))
    covered = {d.year for d in holidays} | {d.year for d in workdays}
    missing = [y for y in years if not ((y in _HOLIDAYS and _HOLIDAYS[y]) or y in covered)]
    return rows, missing


def builtin_holiday_year_range():
    """返回内置节假日数据收录的年份范围（(最小年, 最大年)），供超年份警告使用。"""
    if _HOLIDAYS:
        return min(_HOLIDAYS), max(_HOLIDAYS)
    return 2020, 2030


def default_name_for(stage_tag, start, end, suffix=""):
    base = "设计计划时间表_{0}_{1}_{2}".format(stage_tag, start.isoformat(), end.isoformat())
    return base + suffix + ".xlsx"


def derive_both_paths(base_path):
    """both 模式下由用户选择/指定的单个路径派生完整版与简洁版两个路径。"""
    p = base_path.strip()
    if p.lower().endswith(".xlsx"):
        p = p[:-5]
    if p.endswith("_完整版"):
        return p + ".xlsx", p[:-4] + "_简洁版.xlsx"
    if p.endswith("_简洁版"):
        return p[:-4] + "_完整版.xlsx", p + ".xlsx"
    return p + "_完整版.xlsx", p + "_简洁版.xlsx"


def open_generated(path):
    """生成成功后打开 Excel：Windows 用 os.startfile，Linux 用 xdg-open。

    v5.0.1：打开前先转绝对路径并校验文件存在，修复打包/不同工作目录下
    相对路径导致点击“是”后无响应的缺陷。
    v5.1：Windows 下增加多重降级（os.startfile → cmd start → explorer.exe），
    任一成功即视为打开成功；全部失败或文件缺失时必定给出可见提示
    （控制台打印 + 置顶 GUI 弹窗），不再静默无反应。
    """
    try:
        path = os.path.abspath(path)
        if not os.path.exists(path):
            raise FileNotFoundError("文件不存在：{0}".format(path))
        if sys.platform.startswith("win"):
            last_err = None
            # 方式1：os.startfile（最贴近资源管理器双击行为）
            try:
                os.startfile(path)
                return True
            except Exception as e:
                last_err = e
            # 方式2：cmd /c start "" "path"（兼容关联程序/占用导致的 startfile 失败）
            try:
                subprocess.Popen(["cmd", "/c", "start", "", path])
                return True
            except Exception as e:
                last_err = e
            # 方式3：explorer.exe path（资源管理器兜底，同样会唤起关联程序）
            try:
                subprocess.Popen(["explorer.exe", path])
                return True
            except Exception as e:
                last_err = e
            raise RuntimeError("os.startfile/cmd/explorer 均无法打开：{0}".format(last_err))
        subprocess.Popen(["xdg-open", path],
                         stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL)
        return True
    except Exception as e:
        msg = "自动打开失败：{0}（文件已生成，可手动打开）".format(e)
        print(msg)
        try:
            import tkinter as tk
            from tkinter import messagebox
            _r = tk.Tk()
            _r.withdraw()
            _r.update()
            try:
                _r.attributes("-topmost", True)
                messagebox.showwarning("打开失败", msg, parent=_r)
            finally:
                _r.destroy()
        except Exception:
            pass
        return False


def gui_available():
    """判断当前环境是否适合使用 GUI 确认弹窗。

    环境变量 DSS_NO_GUI=1 可强制禁用（用于自动化/管道测试场景）；
    tkinter 不可用或为非桌面环境时返回 False，调用方回退为命令行 input()。
    """
    if os.environ.get("DSS_NO_GUI", "").strip().lower() in ("1", "yes", "true", "on"):
        return False
    try:
        import tkinter  # noqa: F401
        return True
    except Exception:
        return False


def ask_yes_no_gui(title, message):
    """GUI 确认弹窗：置顶 + 立即刷新事件 + 模态事件循环，点击立即返回。

    使用独立的 Tk 根窗口并 withdraw 隐藏，仅显示 messagebox 对话框；
    弹窗前调用 lift()/update() 强制置顶并处理未决事件，保证弹窗可见且
    立即获得输入焦点，彻底避免“控制台窗口无焦点/被遮挡导致键盘输入
    无效等待”。模态循环内点击“是/否”即刻返回。
    返回 True/False；GUI 不可用或弹窗异常返回 None（由调用方回退 CLI）。
    """
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)  # 置顶，避免被其它窗口遮挡
            root.lift()
            root.update_idletasks()
            root.update()                      # 处理未决事件，弹窗立即可见
            ans = messagebox.askyesno(title, message, parent=root,
                                      icon="question", default="yes")
            return bool(ans)
        finally:
            root.destroy()
    except Exception:
        return None


def ask_open_generated():
    """询问是否立即打开生成的 Excel：GUI 弹窗优先，纯 CLI 回退 input。

    真实桌面（双击 exe / 本机运行）走 tkinter 确认弹窗：置顶聚焦、
    点击立即响应，不依赖控制台输入，消除 CLI/GUI 混用造成的输入延时；
    纯命令行/无显示环境/设置 DSS_NO_GUI=1 时回退为标准 input()，
    保证管道与自动化场景行为与 v3.5 一致。
    """
    if gui_available():
        ans = ask_yes_no_gui("设计计划时间表生成器",
                             "Excel 已生成完毕。\n\n是否立即打开生成的 Excel？")
        if ans is not None:
            return ans
    try:
        s = input("是否立即打开生成的 Excel？(Y/N)：").strip().lower()
        return s in ("y", "yes", "1", "true")
    except EOFError:
        return False


def gui_result_summary(results, start=None, end=None, drawing_days=None):
    """v5.3.2(U2)：GUI 生成完成后弹出结果摘要，再询问是否打开生成文件。

    results: [(fmt, path, node_count, overlap, stats, warnings), ...]
    返回 True=打开文件 / False=不打开 / None=无 GUI 环境（已打印摘要）。
    """
    lines = ["生成完成，共 {0} 份文件：".format(len(results))]
    for f, path, node_count, overlap, stats, warnings in results:
        kind = "完整版" if f == "full" else "简洁版"
        lines.append("")
        lines.append("【{0}】{1}".format(kind, path))
        lines.append("  节点总数：{0} 个".format(node_count))
        if stats:
            _draw = stats.get("drawing_date")
            lines.append("  里程碑：{0} 个　　计划出图日期：{1}".format(
                stats.get("milestone_total", 0),
                fmt_date_cn(_draw) if _draw else "——"))
        for w in warnings:
            lines.append("  提示：{0}".format(w))
        if overlap > 0:
            lines.append("  ⚠ 重合警告：{0} 个日期存在两个及以上时间设定不同的节点重合，需人工协调".format(overlap))
    if start is not None and end is not None and drawing_days is not None:
        lines.append("")
        lines.append("项目周期：{0} 天　　出图制作时间：{1} 天".format(
            (end - start).days + 1, drawing_days))
    if not gui_available():
        print("\n".join(lines))
        return None
    import tkinter.messagebox as _mb
    return _mb.askyesno("生成结果摘要", "\n".join(lines) + "\n\n是否打开生成文件？")


def main(argv=None):
    args = parse_args(argv if argv is not None else sys.argv[1:])
    if args.version:
        print("设计计划时间表生成器 v4.5中国院版（作者：JaJa+AI）")
        return 0
    # v5.0：首次运行建立运行支持文件夹（标准模板/覆写节点文件/应用标准文件）
    try:
        ensure_runtime_dirs()
        ensure_template_samples()  # v5.3.3(U10)：自动内置模板样例
    except Exception as e:
        print("警告：无法建立运行支持文件夹：{0}".format(e))
    cfg = {} if args.no_config else load_config()

    # v4.0 模式判定：完整命令行参数 = CLI 模式；--cli 或 DSS_NO_GUI 强制 CLI 交互；
    # 其余情况默认启动 GUI 图形界面（GUI 不可用时回退 CLI 交互）。
    cli_mode = bool(args.stage and args.start
                    and (args.end or (args.period_days is not None and args.period_days > 0)))
    gui_mode = (not cli_mode) and not args.cli and gui_available()
    interactive = (not cli_mode) and not gui_mode

    # C01：CLI 显式参数预校验——无效参数直接报错返回，不降级交互 input()
    # （避免无 TTY 子进程 EOFError 崩溃 / 有 TTY 时挂起等待）
    stage_key = parse_stage(args.stage) if args.stage else None
    start = parse_date_arg(args.start) if args.start else None
    end = parse_date_arg(args.end) if args.end else None
    if args.stage and stage_key is None:
        print("错误：无法识别的设计阶段“{0}”。可用值：1 / 初步 / preliminary（初步设计），"
              "2 / 施工图 / construction（施工图设计）。".format(args.stage))
        return 1
    if args.start and start is None:
        print("错误：开始日期“{0}”不是有效日期，格式应为 YYYY-MM-DD。".format(args.start))
        return 1
    if args.end and end is None:
        print("错误：结束日期“{0}”不是有效日期，格式应为 YYYY-MM-DD。".format(args.end))
        return 1
    if (args.stage or args.start or args.end or args.period_days is not None) \
            and not cli_mode and not args.cli and not gui_mode:
        print("错误：命令行参数不完整，无法以命令行模式生成。")
        print("用法示例：--stage 1 --start 2026-01-05 --end 2026-02-28（方式1）；"
              "或 --stage 1 --start 2026-01-05 --period-days 55（方式2）。")
        print("提示：当前环境无法启动图形界面，请补齐 --start 与 --end/--period-days 后重试。")
        return 1

    # v4.0：GUI 图形界面模式收集全部参数（取消/关闭窗口返回 None）
    gui_opts = None
    if gui_mode:
        gui_opts = run_gui(args, cfg)
        if gui_opts is None:
            print("已取消生成。")
            return 0
        stage_key = gui_opts["stage_key"]
        start = gui_opts["start"]
        end = gui_opts["end"]

    # 交互补全
    while stage_key is None:
        s = input("请输入设计阶段（1=初步设计阶段，2=施工图设计阶段）：").strip()
        stage_key = parse_stage(s)
        if stage_key is None:
            print("无法识别，请重新输入。")
    while start is None:
        s = input("请输入设计阶段开始日期（格式 YYYY-MM-DD）：").strip()
        start = parse_date_arg(s)
        if start is None:
            print("日期格式错误，请重新输入。")
    if end is None and args.period_days is not None and args.period_days > 0:
        # 方式2（CLI --period-days）：periodDays>0 时忽略 end，自动推算
        end = start + timedelta(days=args.period_days - 1)
        print("已按设计计划周期计算结束日期：{0}（周期 {1} 天，含首尾）。".format(
            end.isoformat(), args.period_days))
    if interactive and end is None:
        # 交互式周期输入方式选择（v3.7）：方式1 开始+结束；方式2 开始+设计总时间（天）
        while True:
            s = input("请选择周期输入方式（1=开始日期+结束日期；2=开始日期+设计计划周期（设计总时间，天））：").strip()
            if s in ("1", "方式1", "开始+结束", "开始日期+结束日期"):
                while True:
                    s2 = input("请输入设计阶段结束日期（格式 YYYY-MM-DD）：").strip()
                    end = parse_date_arg(s2)
                    if end is None:
                        print("日期格式错误，请重新输入。")
                    else:
                        break
                break
            if s in ("2", "方式2", "开始+周期", "开始日期+设计计划周期", "开始日期+设计总时间"):
                while True:
                    s2 = input("请输入设计计划周期（设计总时间，天）：").strip()
                    try:
                        pd = int(s2)
                    except ValueError:
                        pd = 0
                    if pd <= 0:
                        print("请输入正整数（天）。")
                        continue
                    end = start + timedelta(days=pd - 1)
                    print("已按设计计划周期计算结束日期：{0}（周期 {1} 天，含首尾）。".format(
                        end.isoformat(), pd))
                    break
                break
            print("无法识别，请输入 1 或 2。")
    if end <= start:
        print("错误：结束日期必须晚于开始日期。")
        return 1

    # ---- 短周期预警（v3.1）：设计周期少于 15 天 ----
    period_days = (end - start).days + 1  # 设计周期含首尾
    if period_days < 15:
        if interactive:
            while True:
                s = input("警告：设计周期仅 {0} 天，节点可能大量重合。是否继续？(Y/N)：".format(period_days)).strip().lower()
                if s in ("y", "yes", "1", "true"):
                    break
                if s in ("n", "no", "0", "false"):
                    print("已取消生成。")
                    return 1
                print("无法识别，请输入 Y 或 N。")
        else:
            print("警告：设计周期仅 {0} 天，节点可能大量重合（命令行模式已默认继续）。".format(period_days))

    # ---- 周期过长预警（v5.3.2/L2）：设计周期超过 730 天 ----
    if period_days > 730:
        if interactive:
            while True:
                s = input("警告：设计周期长达 {0} 天（超过两年），节点分布将非常稀疏。是否继续？(Y/N)：".format(period_days)).strip().lower()
                if s in ("y", "yes", "1", "true"):
                    break
                if s in ("n", "no", "0", "false"):
                    print("已取消生成。")
                    return 1
                print("无法识别，请输入 Y 或 N。")
        else:
            print("警告：设计周期长达 {0} 天（超过两年），节点分布将非常稀疏（命令行模式已默认继续）。".format(period_days))

    # 项目名称 / 设计编号（可沿用 config.json 记忆；GUI 模式优先取界面输入）
    project_name = args.project if args.project else ""
    design_code = args.code if args.code else ""
    if gui_opts:
        project_name = gui_opts["project_name"] or project_name
        design_code = gui_opts["design_code"] or design_code
    if not project_name:
        if interactive:
            if cfg.get("project_name"):
                s = input("请输入项目名称（上次：{0}，直接回车沿用；留空+回车=不填）：".format(cfg["project_name"])).strip()
                project_name = s if s else cfg["project_name"]
            else:
                project_name = input("请输入项目名称（可留空）：").strip()
        else:
            project_name = ""  # CLI 模式未提供项目名则留空
    if not design_code:
        if interactive:
            if cfg.get("design_code"):
                s = input("请输入设计编号（上次：{0}，直接回车沿用；留空+回车=不填）：".format(cfg["design_code"])).strip()
                design_code = s if s else cfg["design_code"]
            else:
                design_code = input("请输入设计编号（可留空）：").strip()
        else:
            design_code = ""  # CLI 模式未提供设计编号则留空

    # 是否避开节假日（可沿用 config.json 记忆）
    avoid_holidays = None
    if args.avoid_holidays is not None:
        avoid_holidays = parse_yes_no(args.avoid_holidays)
    elif "avoid_holidays" in cfg:
        avoid_holidays = bool(cfg["avoid_holidays"])
    if gui_opts:
        avoid_holidays = bool(gui_opts["avoid_holidays"])
    if interactive:
        default_flag = "Y" if avoid_holidays else "N"
        while True:
            if avoid_holidays is not None:
                s = input("是否避开节假日（将落在休息日的节点迁移到最近工作日并注明原日期）？"
                          "（Y/N，上次：{0}，直接回车沿用）：".format("是" if avoid_holidays else "否")).strip().lower()
                if not s:
                    break
            else:
                s = input("是否避开节假日（将落在休息日的节点迁移到最近工作日并注明原日期）？(Y/N)：").strip().lower()
            if s in ("y", "yes", "1", "true"):
                avoid_holidays = True
                break
            if s in ("n", "no", "0", "false"):
                avoid_holidays = False
                break
            print("无法识别，请输入 Y 或 N。")
    else:
        if avoid_holidays is None:
            avoid_holidays = False  # CLI 模式未指定且无记忆时默认不避开

    # 输出版本（可沿用 config.json 记忆）
    fmt = args.format
    if fmt is None and "format" in cfg and cfg["format"] in ("full", "brief", "both"):
        fmt = cfg["format"]
    if gui_opts:
        fmt = gui_opts["fmt"]
    if interactive:
        default_fmt_text = {
            "full": "完整版", "brief": "简洁版", "both": "完整+简洁"
        }.get(fmt, "完整版")
        while True:
            s = input("输出版本（1=完整版，2=简洁版，3=同时生成完整版与简洁版，"
                      "上次：{0}，直接回车沿用）：".format(default_fmt_text)).strip()
            if not s and fmt is not None:
                break
            if not s:
                fmt = "full"
                break
            parsed = parse_format(s)
            if parsed is not None:
                fmt = parsed
                break
            print("无法识别，请输入 1 / 2 / 3。")
    else:
        if fmt is None:
            fmt = "full"  # CLI 模式未指定且无记忆时默认完整版

    # 出图制作时间（天，可沿用 config.json 记忆）
    drawing_days = args.drawing_days
    if drawing_days is None and "drawing_days" in cfg:
        try:
            drawing_days = int(cfg["drawing_days"])
        except (TypeError, ValueError):
            drawing_days = None
    if gui_opts:
        drawing_days = gui_opts["drawing_days"]
    if interactive:
        default_dd = drawing_days if drawing_days is not None else 1
        while True:
            s = input("出图制作时间（天，上次：{0}，直接回车沿用；默认 1）：".format(default_dd)).strip()
            if not s:
                drawing_days = default_dd
                break
            try:
                drawing_days = int(s)
                if drawing_days < 0:
                    print("天数不能为负，请输入非负整数（0 表示不启用）。")
                    continue
                break
            except ValueError:
                print("请输入整数（天）。")
    else:
        if drawing_days is None:
            drawing_days = 1  # CLI 模式未指定且无记忆时默认 1 天
        # C02：出图制作时间负数不再静默钳制为 0，直接报错返回
        if drawing_days < 0:
            print("错误：出图制作时间不能为负数（当前 {0} 天）。"
                  "请使用 0 表示不启用，或正整数（1~周期-1 天）。".format(drawing_days))
            return 1

    # v5.3.2(A2)：出图制作时间上限校验（CLI/交互/GUI 统一入口，拒绝并提示合理范围；
    # GUI 提交回调亦会弹窗拦截，此处兜底）
    dd_err = validate_drawing_days(drawing_days, start, end)
    if dd_err:
        print("错误：{0}".format(dd_err))
        return 1

    # v5.3.2(A4)：取整模式（--round-mode 优先，其次 config.json 记忆，GUI 覆盖，缺省 round）
    round_mode = args.round_mode
    if round_mode is None and "round_mode" in cfg and cfg["round_mode"] in ("floor", "ceil", "round"):
        round_mode = cfg["round_mode"]
    if round_mode is None:
        round_mode = "round"
    if gui_opts:
        round_mode = gui_opts.get("round_mode", round_mode)

    # v4.0：输出表格风格（--style 优先，其次 config 记忆，GUI 覆盖，缺省 classic）
    _pals = load_style_palettes()  # v5.3.3(U1)：内置+自定义样式合并表
    style = args.style
    if style is None and "style" in cfg and cfg["style"] in _pals:
        style = cfg["style"]
    if style is None or style not in _pals:
        style = "classic"
    if gui_opts:
        style = gui_opts["style"]

    # v4.1：是否包含专业作业清单（--work-lists/--no-work-lists 优先，其次
    # config.json 记忆，GUI 复选框覆盖；缺省不含）
    work_lists = args.work_lists
    if work_lists is None and "work_lists" in cfg:
        work_lists = bool(cfg["work_lists"])
    if work_lists is None:
        work_lists = False
    if gui_opts:
        work_lists = bool(gui_opts.get("work_lists", False))

    # v5.0：多语种选项（--languages 优先，其次 config.json 记忆，GUI 复选框覆盖；
    # 无选择时仅中文）。代码经 normalize_languages 过滤，非法代码自动忽略。
    languages = args.languages if args.languages else []
    if isinstance(languages, str):
        languages = [x.strip() for x in languages.split(",") if x.strip()]
    # C03：显式传入未知语言代码时打印警告（不再静默忽略）
    if args.languages:
        _unknown = sorted({x.strip().lower() for x in languages
                           if x.strip() and x.strip().lower() not in LANG_OPTIONS})
        if _unknown:
            print("警告：无法识别的语言代码：{0}，已忽略。可用代码：{1}。".format(
                "、".join(_unknown),
                "、".join("{0}={1}".format(k, LANG_OPTIONS[k]) for k in LANG_OPTIONS)))
    languages = normalize_languages(languages)
    if not languages and "languages" in cfg:
        languages = normalize_languages(cfg.get("languages", []))
    if gui_opts:
        languages = normalize_languages(gui_opts.get("languages", []))
    if interactive and gui_opts is None:
        # 交互模式询问多语种（列出可用语言，可输入多个代码；回车沿用/无记忆则仅中文）
        default_langs = "、".join(LANG_OPTIONS[c] for c in languages) if languages else "无（仅中文）"
        s = input("多语种选项（可选语言代码：en=英语,ja=日语,ko=韩语,fr=法语,de=德语,es=西班牙语,"
                  "pt=葡萄牙语,ru=俄语,ar=阿拉伯语,it=意大利语,vi=越南语,th=泰语；多个用英文逗号分隔，"
                  "如 en,ja；直接回车沿用上次：{0}）：".format(default_langs)).strip()
        if s:
            _raw_langs = [x.strip() for x in s.split(",") if x.strip()]
            # C03：交互输入未知语言代码时同样警告（不再静默忽略）
            _unknown = sorted({x.lower() for x in _raw_langs if x.lower() not in LANG_OPTIONS})
            if _unknown:
                print("警告：无法识别的语言代码：{0}，已忽略。可用代码：{1}。".format(
                    "、".join(_unknown),
                    "、".join("{0}={1}".format(k, LANG_OPTIONS[k]) for k in LANG_OPTIONS)))
            languages = normalize_languages(_raw_langs)
    # v5.0：覆写节点
    overwrite_nodes = []
    overwrite_extra_profs = []
    if gui_opts:
        for ow in gui_opts.get("overwrite_nodes", []):
            try:
                _src = ow.get("name") or os.path.basename(ow.get("path") or "")
                for _nd in ow["nodes"]:
                    # v5.3：覆写节点名称显示为「模板文件名：模板内节点名」
                    overwrite_nodes.append((_nd[0], _nd[1], _src + "：" + _nd[2]) + tuple(_nd[3:]))
            except Exception:
                pass
            for p in (ow.get("extra_profs") or []):
                if p not in overwrite_extra_profs:
                    overwrite_extra_profs.append(p)
    if args.overwrite_file:
        for fp in args.overwrite_file:
            # C04①：覆写文件不存在时明确报错返回（不再仅一行“忽略”提示后继续生成）
            if not os.path.isfile(fp):
                print("错误：覆写节点模板文件不存在或不是文件：{0}（请检查路径后再运行）。".format(fp))
                return 1
            try:
                ov = parse_overwrite_xlsx(fp, stage_key)
                for _nd in ov["nodes"]:
                    # v5.3：覆写节点名称显示为「模板文件名：模板内节点名」
                    overwrite_nodes.append((_nd[0], _nd[1], ov["name"] + "：" + _nd[2]) + tuple(_nd[3:]))
                for p in (ov.get("extra_profs") or []):
                    if p not in overwrite_extra_profs:
                        overwrite_extra_profs.append(p)
                print("已加载覆写节点模板：{0}（{1} 个节点，内容保持原文不翻译）".format(
                    os.path.basename(fp), len(ov["nodes"])))
                if ov.get("extra_profs"):
                    print("  识别到新专业：{0}，将自动生成对应工作表与作业清单。".format(
                        "、".join(ov["extra_profs"])))
            except Exception as e:
                # C04①：解析失败改为醒目警告（文件存在但内容不符合模板结构）
                print("警告：覆写节点模板解析失败，已忽略该文件：{0}。原因：{1}。"
                      "请确认其为有效的覆写节点模板 xlsx。".format(fp, e))
    # 多语种输出：表格在中文基础上追加所选语言译文，未覆盖词条保持中文。
    if languages:
        print("多语种译文：{0}（表格将在中文基础上追加所选语言译文，词典未覆盖词条保持中文）。".format(
            "、".join(LANG_OPTIONS[c] for c in languages)))

    # v5.3.2(A3)：节假日数据外置——优先读取运行支持目录 holidays.json（允许覆盖内置）；
    # 未指定 --holidays-file 且运行支持目录存在 holidays.json 时自动加载。
    holidays_file = args.holidays_file
    if not holidays_file:
        _rt_hf = os.path.join(program_dir(), RUNTIME_ROOT_NAME, "holidays.json")
        if os.path.isfile(_rt_hf):
            holidays_file = _rt_hf
            print("已自动加载运行支持目录节假日文件：{0}".format(_rt_hf))
    holidays, workdays = load_holidays(holidays_file)
    # v5.3.2(A3)：项目起止日期跨越内置数据未收录年份时输出警告
    ymin, ymax = builtin_holiday_year_range()
    if start.year < ymin or end.year > ymax:
        print("警告：项目起止日期（{0} ~ {1}）超出内置节假日数据收录范围（{2}~{3} 年），"
              "相关年份节假日/调休数据可能不完整，请留意计划日期合理性"
              "（可通过运行支持目录 holidays.json 补充）。".format(
                  start.isoformat(), end.isoformat(), ymin, ymax))

    stage_tag = "初步设计" if stage_key == "preliminary" else "施工图设计"
    suffix = "" if fmt == "full" else ("_简洁版" if fmt == "brief" else "_完整版")

    # 输出路径：命令行 --output 优先；GUI 界面已选路径其次；否则弹出系统保存对话框
    gui_used = False
    if args.output:
        base_out = args.output
    elif gui_opts and gui_opts.get("output_path"):
        base_out = gui_opts["output_path"]
        gui_used = True
    else:
        default_name = default_name_for(stage_tag, start, end, suffix)
        if not interactive and not gui_mode:
            # C04②：纯 CLI 无人值守模式不弹系统保存对话框，直接写入当前目录
            base_out = os.path.join(os.getcwd(), default_name)
            print("未指定输出路径，已直接保存到当前目录：{0}".format(os.getcwd()))
        else:
            chosen, gui_used = choose_output_path(default_name)
            if chosen is None:
                print("已取消保存，未生成文件。可重新运行并用 --output 指定路径。")
                return 1
            base_out = chosen

    if fmt == "both":
        full_path, brief_path = derive_both_paths(base_out)
        targets = [("full", full_path), ("brief", brief_path)]
    else:
        targets = [(fmt, base_out)]

    # v3.6：生成开始前给出进度提示，避免用户在渲染/写盘期间误以为卡死
    print("正在生成 Excel（含样式渲染与写盘），请稍候…")
    results = []
    for f, path in targets:
        try:
            path, node_count, overlap, stats, warnings = generate_excel(
                stage_key, start, end, path, holidays, workdays,
                project_name=project_name, design_code=design_code,
                avoid_holidays=avoid_holidays, fmt=f, drawing_days=drawing_days,
                style=style, work_lists=work_lists,
                languages=languages,
                overwrite_nodes=overwrite_nodes,
                overwrite_extra_profs=overwrite_extra_profs,
                round_mode=round_mode)
        except Exception as e:
            print("生成失败：{0}".format(e))
            return 1
        results.append((f, path, node_count, overlap, stats, warnings))

    # 写回配置记忆（v3.1）
    if not args.no_config:
        new_cfg = {
            "project_name": project_name,
            "design_code": design_code,
            "drawing_days": drawing_days,
            "avoid_holidays": avoid_holidays,
            "format": fmt,
            "style": style,
            "work_lists": work_lists,
            "languages": languages,
            "round_mode": round_mode,
            # v5.3.2(U4)：记忆设计阶段与起止日期，GUI 启动时预填
            "stage_key": stage_key,
            "start_date": start.isoformat(),
            "end_date": end.isoformat(),
        }
        save_config(new_cfg)
        print("配置记忆已写入：{0}".format(CONFIG_PATH))

    for f, path, node_count, overlap, stats, warnings in results:
        kind = "完整版" if f == "full" else "简洁版"
        print("已生成（{0}）：{1}".format(kind, path))
        for w in warnings:
            print("  提示：{0}".format(w))
        print("  项目名称：{0}　　设计编号：{1}".format(
            project_name if project_name else "——", design_code if design_code else "——"))
        print("  节假日处理：{0}　　输出版本：{1}".format(
            "避开（节点迁移至最近工作日）" if avoid_holidays else "不避开", kind))
        print("  出图制作时间：{0} 天　　表格风格：{1}（{2}）　　共排布 {3} 天（{4} 周），自动填入 {5} 个节点日期。".format(
            drawing_days, style, _pals.get(style, _pals["classic"])["desc"],
            stats["total_days"], stats["total_weeks"], node_count))
        print("  项目概况：节点总数 {0} 个，里程碑 {1} 个，工作日 {2} 天，休息日 {3} 天，计划出图日期 {4}。".format(
            stats["node_total"], stats["milestone_total"], stats["work_days"],
            stats["rest_days"], fmt_date_cn(stats["drawing_date"])))
        print("  已生成“节点清单”汇总工作表，便于管理者核对与打印。")
        if work_lists:
            print("  已生成「专业作业清单」工作表：按专业列出提图/收图时间与内容"
                  "（发图行淡橙 #FCE4D6、收图行淡蓝 #DDEBF7）。")
        if languages:
            print("  多语种译文：已在中文基础上追加 {0} 语言译文"
                  "（内置词典未覆盖的词条保持中文）。".format(
                      "、".join(LANG_OPTIONS[c] for c in languages)))
        if overlap > 0:
            print("  ⚠ 检测到 {0} 个日期存在两个及以上时间设定（百分比）不同的节点重合，"
                  "相关日期已标浅红色并提示需人工协调（文档中同百分比多节点同日属正常设计，不计入）。".format(overlap))
    if gui_used:
        print("已通过系统保存对话框选择保存位置。")

    # 生成后自动打开（v3.1）
    do_open = None
    if args.open is not None:
        do_open = parse_yes_no(args.open)
    elif interactive or gui_mode:
        if gui_mode:
            # v5.3.2(U2)：GUI 生成完成后先弹出结果摘要，再询问是否打开生成文件
            do_open = gui_result_summary(results, start, end, drawing_days)
        else:
            # v3.6：GUI 确认弹窗优先（置顶聚焦、点击立即响应，消除 CLI/GUI 混用
            # 导致的输入无效等待）；纯命令行/无显示环境/设置 DSS_NO_GUI=1 时回退 input()
            do_open = ask_open_generated()
    if do_open:
        for f, path, _, _, _, _ in results:
            open_generated(path)
    elif not interactive and not gui_mode:
        pass  # CLI 模式默认不自动打开（除非 --open yes）
    return 0


if __name__ == "__main__":
    sys.exit(main())
