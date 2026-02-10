# Lawyor 项目说明文档 (Project Documentation)

> **注：** 本文档包含图片占位符。由于环境限制无法直接生成图片文件，请根据占位符提示补充相应的图片和截图。

---

### 1. 愿景与定位 (Vision & Positioning)

**Slogan:**
> **Lawyor**
> **From Lawyer to Creator.** (从律师到创造者)
> **Cursor for Legal Professionals.** (法律人的 AI 智能工作台)

**[图片占位符 1：产品愿景概念图]**
> **图片描述与生成指南 (Image Prompt):**
> *   **核心画面**：一张充满未来感的抽象概念图，表现“法律文档”与“数字代码”的深度融合。
> *   **视觉焦点**：画面中央是一个古老的、发光的羊皮卷轴（象征传统法律合同），它正在向右侧延展并发生“数字化质变”，逐渐转化为蓝白色的二进制数据流（010101）和结构化的程序代码块（Code Blocks）。
> *   **风格基调**：Professional (专业), High-tech (高科技), Minimalist (极简主义)。
> *   **色彩搭配**：背景为深邃的科技蓝（Deep Blue），配合白色的光线和少许金色的点缀（代表法律的庄严）。
> *   **光影效果**：电影级布光（Cinematic Lighting），强调从左至右的进化感。
> *   **生成提示词参考 (Midjourney/DALL-E)**: "A futuristic abstract representation of a legal contract document merging with digital computer code. A glowing parchment scroll in the center transforming into streams of blue and white binary data and structured code blocks. Professional, high-tech, minimalist style. Cinematic lighting, deep blue background with gold accents. 8k resolution, unreal engine 5 render. --ar 16:9"

**产品简介 (Introduction):**
**Lawyor** 是一款专为法律专业人士打造的跨平台智能文档 IDE（集成开发环境），完美适配 **WPS Office** 与 **Microsoft Word**。

我们致力于将代码编辑器（如 Cursor）的高效理念引入法律领域。在 **Lawyor** 的辅助下，法律文档不再是枯燥的文字堆砌，而是结构化的逻辑产物。通过 AI 驱动的“内容-格式”分离架构，Lawyor 让律师从机械的填写、校对和排版工作中彻底解放，只专注于法律条款的推敲与商业逻辑的思考。

---

### 2. 品牌理念 (Brand Philosophy)

**Why "Lawyor"?**
我们将 "Lawyer"（律师）的后缀升级为 "-or"，象征着 **Executor（执行者）** 与 **Creator（创造者）** 的结合。

在传统模式下，律师往往被繁琐的文档格式所束缚，是被动的“处理者”。而在 **Lawyor** 的辅助下，律师将进化为法律逻辑的**创造者**——像编写代码一样构建合同架构，像运行程序一样自动生成文档。Lawyor 不仅是工具，更是法律人思维的外脑。

---

### 3. 核心痛点 (Core Pain Points)

传统法律工作中，律师面临着“三大枷锁”：

*   **重复劳动 (Repetitive Labor):** 在处理同类合同时，反复查找替换主体名称、金额、日期等变量，不仅效率低下，而且极易因人工疏忽导致遗漏或错误。
*   **格式噩梦 (Formatting Nightmare):** 调整缩进、字体、编号层级耗时耗力，往往牵一发而动全身。律师的大量精力被迫消耗在“对齐空格”上，而非法律分析上。
*   **知识复用难 (Knowledge Silo):** 资深律师的审查经验（Checklist）难以固化为可执行的工具，新人上手慢，团队协作标准难以统一。

---

### 4. 解决方案与核心功能 (Solutions & Core Features)

**Lawyor** 通过 "Hollow Out"（挖空）与 "Fill In"（注入）的双向智能流，重新定义了合同生产流程：

#### 4.1 智能模版化：挖空 (Intelligent Templating: Hollow Out)
当一份新合同导入时，Lawyor 的 AI 引擎会自动扫描全文，识别出“公司名称”、“金额”、“日期”、“特殊条款”等所有变量。它会自动在文档中建立 **Content Control（内容控件）** 锚点，将一篇非结构化的 Word 文档瞬间转化为一份结构化的**动态表单**。

**[🔴 待补充截图 A：智能挖空/表单填写]**
> **截图说明：**
> *   **场景**：打开一份合同，侧边栏（Taskpane）展开。
> *   **内容**：侧边栏显示生成的表单字段（如公司名、日期、金额），文档正文中对应的文字有蓝色的框（Content Control）。
> *   **用途**：直观展示“内容结构化”效果。

#### 4.2 结构化起草：审查清单与技能 (Structured Drafting: Checklist & Skills)
我们将资深律师的经验凝练为 **Skill（技能）** 和 **Checklist（审查清单）** 并内置于插件中。
*   **知识注入**：用户上传合同或知识库，AI 自动提炼出针对该类型合同的审查要点。
*   **智能联动**：当用户在表单中勾选选项（如“是否包含对赌条款”）时，Lawyor 会根据审查清单自动插入、隐藏或修改对应的法律条款。用户只需做选择题，无需手动逐字起草。

**[🔴 待补充截图 B：审查清单/校验 (若已开发UI)]**
> **截图说明：**
> *   **场景**：侧边栏显示校验结果或风险提示。
> *   **内容**：展示一个填写错误或未填写的字段被高亮（红色），或者侧边栏显示“风险提示/Checklist”。
> *   **用途**：展示智能风控能力。

#### 4.3 AI 自然语言排版 (AI Natural Language Formatting)
Lawyor 摒弃了复杂的格式菜单，引入了对话式排版系统。
*   **所想即所得**：用户只需在侧边栏聊天框输入指令——“*把一级标题改成黑体三号居中，正文宋体小四，行距固定28磅*”，AI 即可精准操控文档格式。
*   **容错机制**：支持“撤销上一步”和“重试”，让排版过程像修改代码一样安全可控。

**[🔴 待补充截图 C：AI 对话排版]**
> **截图说明：**
> *   **场景**：切换到“AI 排版”或“Chat”标签页。
> *   **内容**：在输入框输入类似“将标题设为黑体三号”，上方显示 AI 的回复“已执行格式修改...”。
> *   **用途**：展示“自然语言交互”能力。

#### 4.4 干净交付 (Clean Delivery)
在合同定稿准备发送给客户前，插件提供“一键清洗”功能。系统会自动移除所有 AI 埋点、辅助标记和隐藏段落，生成一份纯净、标准的 .docx 文档。

**[🔴 待补充截图 D：交付清理]**
> **截图说明：**
> *   **场景**：点击“合同交付”按钮后的弹窗。
> *   **内容**：展示进度条面板（Check -> Backup -> Cleanup）。
> *   **用途**：展示“专业交付”流程。

---

### 5. 交互设计 (Interaction Design)

**Lawyor** 采用 **"Sidebar-First"（侧边栏优先）** 的沉浸式交互设计，确保律师的视线始终停留在文档正文，而不被弹窗打断。

*   **侧边栏导航 (Sidebar Navigation):** 所有的表单填写、变量管理均在右侧 Taskpane（任务窗格）完成。左侧看条款，右侧填数据，左右对照，一目了然。
*   **对话式交互 (Conversational UI):** 在“AI 排版”模式下，交互转变为 Chatbot 形式。用户像在这个聊天窗口和我对话一样，告诉插件“*把所有金额加粗*”，操作即刻完成。
*   **非侵入式提示 (Non-intrusive Notification):** 系统状态、错误提示均通过顶部或底部的轻量级 Toast 消息展示，告别传统插件烦人的 `alert()` 弹窗。

---

### 6. 技术架构特色 (Technical Architecture Highlights)

**[图片占位符 2：技术架构示意图]**
> **图片描述与生成指南 (Image Prompt):**
> *   **核心画面**：一张干净、专业的等距（Isometric/2.5D）技术架构图，背景为纯白或极淡的灰色。
> *   **架构层级（由左至右）**：
>     1.  **左侧 (Input Layer)**：悬浮的 3D 图标，分别代表 **Microsoft Word**（蓝色文档）和 **WPS Office**（橙红色文档），象征双平台适配。
>     2.  **中间 (Processing Layer)**：一个充满科技感的、半透明的玻璃棱镜或芯片图标，代表 **Lawyor Plugin Engine**（核心处理引擎）。从左侧文档发出的光线汇聚于此。
>     3.  **右侧 (Intelligence Layer)**：一个发光的、由节点组成的数字大脑或云端网络，代表 **AI Model**（多模态大模型）。
> *   **数据流向**：使用流畅的线条或光束连接这三个部分，表现数据从文档流向引擎，再流向 AI，最后处理后的结构化数据流回文档的闭环过程。
> *   **风格基调**：Vector Illustration (矢量插画), Clean (干净), Business Technology (商务科技)。
> *   **色彩搭配**：以科技蓝（Blue）、银灰（Silver）为主，配以少许橙色（Orange）作为点缀。
> *   **生成提示词参考 (Midjourney/DALL-E)**: "An isometric technical diagram on a clean white background. Three distinct layers connected by data flow lines. Layer 1 (Left): 3D icons of Microsoft Word and WPS Office documents. Layer 2 (Center): A sleek, futuristic glass prism representing the Lawyor plugin engine. Layer 3 (Right): A glowing digital brain network representing the AI model. Clean vector illustration style, business technology aesthetic, blue and silver color palette. --ar 16:9"

#### 6.1 双平台原生集成 (Dual-Platform Native Integration)
基于标准化的 **Office JavaScript API** 构建核心逻辑，Lawyor 实现了一套代码同时原生适配 **WPS Office** 和 **Microsoft Word**。无论是外企常用的 Word 环境，还是政企常用的 WPS 环境，都能提供一致的流畅体验。

#### 6.2 深度文档控制 (Deep Document Control)
不同于市面上简单的文本替换插件，Lawyor 利用底层的 **Content Control (内容控件)** 技术实现了数据与文档的**实时双向绑定**。
*   **表单变，文档变；文档变，表单变。**
*   **稳健的 Undo/Restore 机制**：我们实现了“非破坏性埋点”，即便用户操作失误，也能通过原生 API 的原子操作完美还原文档的原始内容和格式。

#### 6.3 数据隐私与文档隔离 (Data Privacy & Document Isolation)
Lawyor 采用 **"Document as Database"（文档即数据库）** 的架构。我们将表单数据、逻辑规则直接加密存储在文档文件本身的 **CustomDocumentProperties（自定义属性）** 中。
*   **离线可用**：不依赖云端数据库，文件发给同事，插件打开即用。
*   **隐私安全**：完美符合法律行业对数据隐私和离线作业的严苛要求。

#### 6.4 多模态 AI 架构 (Multi-Model AI Architecture)
*   **灵活接入**：后端架构支持热切换 AI 模型（Claude/Kimi/DeepSeek）。
*   **本地化处理**：对于极度敏感的合同，支持通过本地部署的小模型进行基础处理，最大程度保障数据安全。

---

### 7. 总结 (Summary)

**Lawyor** 不仅仅是一个插件，它是法律文档生产力的**倍增器**。通过将 **WPS/Word** 升级为 **Legal IDE**，我们让律师从繁琐的格式工作中解放出来，真正实现 **"Lawyer be Creator"** —— 让每一位法律人都能专注于创造最核心的法律价值。
