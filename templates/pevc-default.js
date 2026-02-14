const PEVC_DEFAULT_TEMPLATE = 
[
  {
    "id": "section_files",
    "header": {
      "label": "1. 所需文件",
      "tag": "Section_Files"
    },
    "fields": []
  },
  {
    "id": "section_company_info",
    "header": {
      "label": "2. 公司基本信息",
      "tag": "Section_CompanyInfo"
    },
    "fields": [
      {
        "id": "signingDate",
        "label": "签订时间",
        "tag": "SigningDate",
        "type": "date",
        "formatFn": "dateUnderline",
        "placeholder": "选择日期",
        "hasParagraphToggle": true
      },
      {
        "id": "signingPlace",
        "label": "签订地点",
        "tag": "SigningPlace",
        "type": "text",
        "placeholder": "如：北京"
      },
      {
        "id": "lawyerRep",
        "label": "律师代表",
        "tag": "LawyerRepresenting",
        "type": "radio",
        "options": ["公司", "投资方", "公司/投资方"]
      },
      {
        "id": "projectShortName",
        "label": "项目简称",
        "tag": "ProjectShortName",
        "type": "text"
      },
      {
        "id": "companyName",
        "label": "目标公司名称",
        "tag": "CompanyName",
        "type": "text"
      },
      {
        "id": "companyBusiness",
        "label": "主营业务",
        "tag": "CompanyBusiness",
        "type": "text"
      },
      {
        "id": "companyCapital",
        "label": "注册资本",
        "tag": "CompanyCapital",
        "type": "text"
      },
      {
        "id": "companyCity",
        "label": "所在城市",
        "tag": "CompanyCity",
        "type": "text"
      },
      {
        "id": "regAddress",
        "label": "注册地址",
        "tag": "RegAddress",
        "type": "text"
      },
      {
        "id": "legalRep",
        "label": "法定代表人姓名",
        "tag": "LegalRepName",
        "type": "text"
      },
      {
        "id": "legalRepTitle",
        "label": "法定代表人职务",
        "tag": "LegalRepTitle",
        "type": "select",
        "options": ["董事长", "执行董事", "总经理"]
      },
      {
        "id": "legalRepNationality",
        "label": "法定代表人国籍",
        "tag": "LegalRepNationality",
        "type": "select",
        "options": ["中国", "美国", "新加坡", "其他"]
      },
      {
        "id": "businessDesc",
        "label": "主营业务描述",
        "tag": "BusinessDesc",
        "type": "text"
      },
      {
        "id": "currentDirectors",
        "label": "现任董事姓名",
        "tag": "CurrentDirectors",
        "type": "text",
        "placeholder": "多个请用逗号隔开"
      },
      {
        "id": "shareholderCount",
        "label": "股东总数",
        "tag": "ShareholderCount",
        "type": "number",
        "value": "1",
        "autoCount": true,
        "placeholder": "系统自动统计，也可手动修改"
      },
      { "type": "divider", "label": "股东 1 (创始人/大股东)" },
      { "id": "sh1_name", "label": "姓名/名称", "tag": "SH1_Name", "type": "text" },
      { "id": "sh1_type", "label": "类型", "tag": "SH1_Type", "type": "select", "options": ["个人", "有限公司", "合伙企业"] },
      { "id": "sh1_id", "label": "证件号码", "tag": "SH1_ID", "type": "text" },
      { "id": "sh1_nation", "label": "国籍/所在地", "tag": "SH1_Nation", "type": "text" },
      { "id": "sh1_address", "label": "注册地址", "tag": "SH1_Address", "type": "text" },
      { "id": "sh1_reg_cap", "label": "认缴注册资本(万元)", "tag": "SH1_RegCapital", "type": "number" },
      { "id": "sh1_paid_cap", "label": "实缴注册资本(万元)", "tag": "SH1_PaidCapital", "type": "number" },
      { "id": "sh1_ratio", "label": "持股比例/出资比例(%)", "tag": "SH1_Ratio", "type": "number" },
      { "id": "sh1_currency", "label": "币种", "tag": "SH1_Currency", "type": "select", "options": ["人民币", "美元"] },
      { "type": "divider", "label": "增资后" },
      { "id": "sh1_post_reg_cap", "label": "增资后注册资本(万元)", "tag": "SH1_PostRegCapital", "type": "number" },
      { "id": "sh1_post_ratio", "label": "增资后持股比例(%)", "tag": "SH1_PostRatio", "type": "number" }
    ]
  },
  {
    "id": "section_existing_shareholders",
    "type": "existing_shareholders",
    "header": {
      "label": "2.1 现有股东/历轮投资人",
      "tag": "Section_ExistingShareholders"
    },
    "shareholders": [
      { "id": "sh2", "label": "创始股东 2", "tag": "SH2" },
      { "id": "sh3", "label": "种子轮投资人 1", "tag": "SH3" },
      { "id": "sh4", "label": "种子轮投资人 2", "tag": "SH4" },
      { "id": "sh5", "label": "天使轮投资人 1", "tag": "SH5" },
      { "id": "sh6", "label": "天使轮投资人 2", "tag": "SH6" },
      { "id": "sh7", "label": "Pre-A轮投资人 1", "tag": "SH7" },
      { "id": "sh8", "label": "Pre-A轮投资人 2", "tag": "SH8" },
      { "id": "sh9", "label": "A轮投资人 1", "tag": "SH9" },
      { "id": "sh10", "label": "A轮投资人 2", "tag": "SH10" },
      { "id": "sh11", "label": "B轮投资人 1", "tag": "SH11" },
      { "id": "sh12", "label": "B轮投资人 2", "tag": "SH12" }
    ],
    "shareholderFields": [
      { "id": "_name", "label": "姓名/名称", "tag": "_Name", "type": "text" },
      { "id": "_short", "label": "简称", "tag": "_Short", "type": "text" },
      { "id": "_round", "label": "融资轮次", "tag": "_Round", "type": "select", "options": ["创始", "种子轮", "天使轮", "Pre-A轮", "A轮", "B轮", "C轮", "其他"] },
      { "id": "_type", "label": "类型", "tag": "_Type", "type": "select", "options": ["个人", "有限公司", "有限合伙"], "triggerConditional": true },
      { "id": "_id", "label": "证件号码", "tag": "_ID", "type": "text" },
      { "id": "_nation", "label": "国籍/所在地", "tag": "_Nation", "type": "text" },
      { "id": "_address", "label": "注册地址", "tag": "_Address", "type": "text" },
      { "id": "_legalRep", "label": "法定代表人", "tag": "_LegalRep", "paraTag": "_LegalRepPara", "type": "text", "showWhen": ["有限公司", "有限合伙"], "hasParagraphToggle": true },
      { "id": "_investAmount", "label": "投资额(万元)", "tag": "_InvestAmount", "type": "number", "showWhenRound": ["种子轮", "天使轮", "Pre-A轮", "A轮", "B轮", "C轮", "其他"] },
      { "id": "_regCapital", "label": "认缴注册资本(万元)", "tag": "_RegCapital", "type": "number" },
      { "id": "_paidCapital", "label": "实缴注册资本(万元)", "tag": "_PaidCapital", "type": "number" },
      { "id": "_ratio", "label": "持股比例/出资比例(%)", "tag": "_Ratio", "type": "number" },
      { "id": "_currency", "label": "币种", "tag": "_Currency", "type": "select", "options": ["人民币", "美元"] },
      { "id": "_postRegCapital", "label": "增资后注册资本(万元)", "tag": "_PostRegCapital", "type": "number" },
      { "id": "_postRatio", "label": "增资后持股比例(%)", "tag": "_PostRatio", "type": "number" }
    ]
  },
  {
    "id": "section_financing",
    "header": { "label": "3. 本轮融资信息", "tag": "Section_Financing" },
    "fields": [
      {
        "id": "needEquityAdjust",
        "label": "增资前是否需要调整股权",
        "tag": "NeedEquityAdjust",
        "type": "radio",
        "options": ["否", "是"],
        "subFields": [
          { "type": "divider", "label": "股权调整事项 1" },
          { "id": "adj1_type", "label": "调整方式", "tag": "Adj1_Type", "type": "select", "options": ["转出", "增资", "减资"] },
          { "id": "adj1_transferor", "label": "出让方/增资方", "tag": "Adj1_Transferor", "type": "text" },
          { "id": "adj1_transferee", "label": "受让方", "tag": "Adj1_Transferee", "type": "text" },
          { "id": "adj1_price", "label": "价格(万元)", "tag": "Adj1_Price", "type": "number" },
          { "type": "divider", "label": "股权调整事项 2" },
          { "id": "adj2_type", "label": "调整方式", "tag": "Adj2_Type", "type": "select", "options": ["转出", "增资", "减资"] },
          { "id": "adj2_transferor", "label": "出让方/增资方", "tag": "Adj2_Transferor", "type": "text" },
          { "id": "adj2_transferee", "label": "受让方", "tag": "Adj2_Transferee", "type": "text" },
          { "id": "adj2_price", "label": "价格(万元)", "tag": "Adj2_Price", "type": "number" }
        ]
      },
      { "type": "divider", "label": "本次增资" },
      { "id": "investmentAmount", "label": "投资款总额(万元)", "tag": "InvestmentAmount", "type": "number" },
      { "id": "capitalIncrease", "label": "计入注册资本(万元)", "tag": "CapitalIncrease", "type": "number" },
      { "id": "capitalReserve", "label": "计入资本公积金", "tag": "CapitalReserve", "type": "text", "value": "剩余部分", "placeholder": "填'剩余部分'或具体数额" },
      { "id": "postCapitalTotal", "label": "增资后总注册资本(万元)", "tag": "PostCapitalTotal", "type": "number" },
      { "id": "newEquityRatio", "label": "本次取得股权比例(%)", "tag": "NewEquityRatio", "type": "number" },
      { "type": "divider", "label": "基础条款" },
      { "id": "paymentDeadline", "label": "最晚缴纳时间", "tag": "PaymentDeadline", "type": "date" }
    ]
  },
  {
    "id": "section_current_investors",
    "type": "current_investors",
    "header": { "label": "3.1 本轮投资人", "tag": "Section_CurrentInvestors" },
    "investors": [
      { "id": "lead", "label": "领投方", "tag": "Inv_Lead" },
      { "id": "follow1", "label": "跟投方 1", "tag": "Inv_Follow1" },
      { "id": "follow2", "label": "跟投方 2", "tag": "Inv_Follow2" },
      { "id": "follow3", "label": "跟投方 3", "tag": "Inv_Follow3" }
    ],
    "investorFields": [
      { "id": "_name", "label": "名称/姓名", "tag": "_Name", "type": "text" },
      { "id": "_short", "label": "简称", "tag": "_Short", "type": "text" },
      { "id": "_type", "label": "类型", "tag": "_Type", "type": "select", "options": ["有限公司", "有限合伙", "个人"], "triggerConditional": true },
      { "id": "_nation", "label": "注册地/国籍", "tag": "_Nation", "type": "text" },
      { "id": "_address", "label": "注册地址", "tag": "_Address", "type": "text" },
      { "id": "_id", "label": "证件号码", "tag": "_ID", "type": "text" },
      { "id": "_legalRep", "label": "法定代表人", "tag": "_LegalRep", "paraTag": "_LegalRepPara", "type": "text", "showWhen": ["有限公司", "有限合伙"], "hasParagraphToggle": true },
      { "id": "_amount", "label": "投资额(万元)", "tag": "_Amount", "type": "number" },
      { "id": "_currency", "label": "币种", "tag": "_Currency", "type": "select", "options": ["人民币", "美元"] },
      { "id": "_equityRatio", "label": "本次取得股权比例(%)", "tag": "_EquityRatio", "type": "number" },
      { "id": "_regCapital", "label": "本次对应注册资本(万元)", "tag": "_RegCapital", "type": "number" },
      { "id": "_postRegCapital", "label": "增资后注册资本(万元)", "tag": "_PostRegCapital", "type": "number" },
      { "id": "_postRatio", "label": "增资后持股比例(%)", "tag": "_PostRatio", "type": "number" }
    ]
  },
  {
    "id": "section_definitions",
    "header": { "label": "4. 定义及其他签约方", "tag": "Section_Definitions" },
    "fields": [
      { "id": "otherParties", "label": "其他签约方信息", "tag": "OtherParties", "type": "text", "placeholder": "如有其他方请在此备注" }
    ]
  },
  {
    "id": "section_board",
    "header": { "label": "5. 创始人、新董事会、核心员工", "tag": "Section_Board" },
    "fields": [
      { "id": "newBoardSize", "label": "新董事会由几名董事组成", "tag": "NewBoardSize", "type": "number" },
      { "id": "investorBoardSeats", "label": "本轮投资方有权任命董事人数", "tag": "InvestorBoardSeats", "type": "number" },
      { "id": "founderBoardSeats", "label": "创始人有权任命董事人数", "tag": "FounderBoardSeats", "type": "number" },
      { "id": "founderHasOutsideEquity", "label": "创始人是否持有集团外公司股权", "tag": "FounderHasOutsideEquity", "type": "radio", "options": ["是", "否"] },
      { "id": "coreStaffList", "label": "核心员工名单 (姓名/职务)", "tag": "CoreStaffList", "type": "text" }
    ]
  },
  {
    "id": "section_indemnity",
    "header": { "label": "6. 特殊赔偿及其他", "tag": "Section_Indemnity" },
    "fields": [
      { "id": "indemnity_social", "label": "1. 社保/公积金未足额缴纳", "tag": "Indemnity_SocialSecurity", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "indemnity_tax", "label": "2. 未足额缴纳税款/滞纳金", "tag": "Indemnity_Tax", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "indemnity_penalty", "label": "3. 行政处罚或责任", "tag": "Indemnity_Penalty", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "indemnity_license", "label": "4. 业务牌照/资质缺失", "tag": "Indemnity_License", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "indemnity_equity", "label": "5. 股权权属纠纷", "tag": "Indemnity_Equity", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "indemnity_ip", "label": "6. 知识产权侵权/权属不完善", "tag": "Indemnity_IP", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "indemnity_litigation", "label": "7. 未决诉讼/仲裁", "tag": "Indemnity_Litigation", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "indemnity_noncompete", "label": "8. 核心员工违反竞业/保密义务", "tag": "Indemnity_NonCompete", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "type": "divider", "label": "责任限制" },
      { "id": "liability_threshold", "label": "免责门槛金额(万元)", "tag": "Liability_Threshold", "type": "number", "placeholder": "如：50", "formatFn": "chineseNumber" },
      { "id": "warranty_valid_years", "label": "声明保证有效期(年)", "tag": "Warranty_ValidYears", "type": "number", "placeholder": "如：4", "formatFn": "chineseNumber" },
      { "type": "divider", "label": "交易费用" },
      { "id": "fee_success", "label": "交易成功 - 公司承担费用", "tag": "Fee_Success", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true, "subFields": [{ "id": "fee_cap", "label": "费用上限金额(万元)", "tag": "FeeCap", "type": "number", "placeholder": "如：50" }] },
      { "id": "fee_fail", "label": "交易终止 - 各方自担费用", "tag": "Fee_Fail", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "arbitrationOrg", "label": "仲裁机构", "tag": "ArbitrationOrg", "type": "text", "value": "中国国际经济贸易仲裁委员会" },
      { "id": "arbitrationPlace", "label": "仲裁地", "tag": "ArbitrationPlace", "type": "text", "value": "北京" },
      { "id": "hasTS", "label": "是否签署投资意向书", "tag": "HasTS", "type": "radio", "options": ["是", "否"] },
      { "id": "tsDate", "label": "意向书签署日期", "tag": "TSDate", "type": "date" }
    ]
  },
  {
    "id": "section_preemptive",
    "header": { "label": "7. 股权变动限制", "tag": "Section_Preemptive" },
    "fields": [
      { "id": "transfer_restricted_party", "label": "被限制转让的主体", "tag": "TransferRestrictedParty", "type": "text", "value": "创始股东", "placeholder": "例如：创始股东、现有股东" },
      { "id": "transfer_consent", "label": "转让股权需经谁同意", "tag": "TransferConsentSubject", "type": "text", "value": "本轮投资方" },
      { "id": "transfer_consent_type", "label": "同意形式", "tag": "TransferConsentType", "type": "text", "value": "书面同意" },
      { "id": "investorTransferRight", "label": "投资人是否可自由转股", "tag": "InvestorTransferRight", "type": "radio", "options": ["是", "否"], "value": "是" },
      { "id": "hasPreemptiveRight", "label": "新股优先认购权", "tag": "HasPreemptiveRight", "type": "radio", "options": ["是", "否"] },
      { "id": "preemptiveHolder", "label": "优先认购权人", "tag": "PreemptiveHolder", "type": "text", "value": "本轮投资方" },
      { "id": "hasSuperPreemptive", "label": "是否享有超额认购权", "tag": "HasSuperPreemptive", "type": "radio", "options": ["是", "否"] },
      { "id": "hasRofr", "label": "老股优先购买权", "tag": "HasRofr", "type": "radio", "options": ["是", "否"] },
      { "id": "hasCoSale", "label": "共同出售权", "tag": "HasCoSale", "type": "radio", "options": ["是", "否"] },
      { "id": "rofrHolder", "label": "权利享有方", "tag": "RofrHolder", "type": "text", "value": "本轮投资方" },
      { "id": "hasDragAlong", "label": "领售权 (拖售权)", "tag": "HasDragAlong", "type": "radio", "options": ["是", "否"] },
      { "id": "dragAlongTrigger", "label": "领售触发条件", "tag": "DragAlongTrigger", "type": "text", "placeholder": "例如：交割后 5 年未上市" },
      { "id": "dragAlongValuation", "label": "领售最低估值 (亿元)", "tag": "DragAlongValuation", "type": "number" }
    ]
  },
  {
    "id": "section_economics",
    "header": { "label": "8. 核心经济条款", "tag": "Section_Economics" },
    "fields": [
      { "id": "antiDilution", "label": "反稀释权条款", "tag": "HasAntiDilution", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "antiDilutionHolder", "label": "反稀释权人", "tag": "AntiDilutionHolder", "type": "text", "value": "本轮投资方" },
      { "id": "antiDilutionOrigPrice", "label": "本轮原始认购价格(元/注册资本)", "tag": "AntiDilutionOrigPrice", "type": "number", "placeholder": "例如：10" },
      { "id": "antiDilutionMethod", "label": "价格调整方式", "tag": "AntiDilutionMethod", "type": "select", "options": ["广义加权平均", "完全棘轮", "狭义加权平均"] },
      { "id": "antiDilutionCompDays", "label": "补偿期限(天)", "tag": "AntiDilutionCompDays", "type": "number", "value": "30", "formatFn": "chineseNumber" },
      { "id": "preemptiveClauseRef", "label": "优先认购权条款编号", "tag": "PreemptiveClauseRef", "type": "text", "placeholder": "例如：第5.1条" },
      { "id": "liquidationPref", "label": "清算优先权", "tag": "HasLiquidationPref", "type": "radio", "options": ["是", "否"] },
      { "id": "liqRanking", "label": "是否优于普通股", "tag": "LiqRanking", "type": "radio", "options": ["是", "否"] },
      { "id": "liqMultiple", "label": "优先清算回报倍数 (X倍本金)", "tag": "LiqMultiple", "type": "number", "value": "1" },
      { "id": "liqInterest", "label": "清算年化利率 (%)", "tag": "LiqInterest", "type": "number", "value": "0" },
      { "id": "participationType", "label": "剩余财产分配方式", "tag": "ParticipationType", "type": "select", "options": ["无参与权(Non-participating)", "完全参与(Full participating)", "附上限参与(Capped)"] }
    ]
  },
  {
    "id": "section_redemption",
    "header": { "label": "8.1 回购权", "tag": "Section_Redemption" },
    "hasSectionToggle": true,
    "fields": [
      { "id": "hasRedemptionRight", "label": "回购权条款", "tag": "Section_Redemption", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "type": "divider", "label": "回购触发事件" },
      { "id": "redemptionEvent_IPO", "label": "事件1: 未上市/退出失败", "tag": "RedemptionEvent_IPO", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true, "subFields": [{ "id": "redemptionTriggerYears", "label": "触发年限(年)", "tag": "RedemptionTriggerYears", "type": "number", "value": "6", "formatFn": "chineseNumber" }] },
      { "id": "redemptionEvent_Breach", "label": "事件2: 严重违反协议", "tag": "RedemptionEvent_Breach", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "redemptionEvent_Law", "label": "事件3: 严重违反法律法规", "tag": "RedemptionEvent_Law", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "redemptionEvent_Policy", "label": "事件4: 法律政策变化", "tag": "RedemptionEvent_Policy", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "redemptionEvent_Founder", "label": "事件5: 创始人/核心人员问题", "tag": "RedemptionEvent_Founder", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "redemptionEvent_Control", "label": "事件6: 实际控制人变更", "tag": "RedemptionEvent_Control", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "redemptionEvent_Business", "label": "事件7: 主营业务变更/经营异常", "tag": "RedemptionEvent_Business", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "type": "divider", "label": "回购主体" },
      { "id": "redemptionRightHolder", "label": "回购权人", "tag": "RedemptionRightHolder", "type": "text", "value": "本轮投资方与投资方" },
      { "id": "redemptionObligor", "label": "回购义务人", "tag": "RedemptionObligor", "type": "text", "value": "公司与创始股东" },
      { "id": "redemptionClauseRef", "label": "回购价格条款编号", "tag": "RedemptionClauseRef", "type": "text", "value": "第3.2条" },
      { "type": "divider", "label": "期限与违约" },
      { "id": "redemptionNotifyDays", "label": "通知其他回购权人期限(工作日)", "tag": "RedemptionNotifyDays", "type": "number", "value": "3", "formatFn": "chineseNumber" },
      { "id": "redemptionPaymentDays", "label": "回购支付期限(日)", "tag": "RedemptionPaymentDays", "type": "number", "value": "40", "formatFn": "chineseNumber" },
      { "id": "redemptionPenaltyRate", "label": "违约金利率(每日万分之)", "tag": "RedemptionPenaltyRate", "type": "number", "value": "5" },
      { "id": "redemptionAssetSaleDays", "label": "资产变卖触发期限(日)", "tag": "RedemptionAssetSaleDays", "type": "number", "value": "90", "formatFn": "chineseNumber" },
      { "type": "divider", "label": "回购顺序" },
      { "id": "redemptionPriorityHolder", "label": "第一顺位(优先支付方)", "tag": "RedemptionPriorityHolder", "type": "text", "value": "本轮投资方" },
      { "id": "redemptionSecondaryHolder", "label": "第二顺位", "tag": "RedemptionSecondaryHolder", "type": "text", "value": "投资方" }
    ]
  },
  {
    "id": "section_other_rights",
    "header": { "label": "9. 其他优先权", "tag": "Section_OtherRights" },
    "fields": [
      { "id": "ipo_auto_convert", "label": "IPO自动转股机制", "tag": "IPOAutoConvert", "type": "radio", "options": ["是", "否"] },
      { "id": "ipo_min_valuation", "label": "合格IPO最低估值 (亿元)", "tag": "IPOMinValuation", "type": "number", "value": "40" },
      { "id": "ipo_min_proceeds", "label": "合格IPO最低募资额 (亿元)", "tag": "IPOMinProceeds", "type": "number", "value": "10" },
      { "id": "hasInfoRights", "label": "信息权", "tag": "HasInfoRights", "type": "radio", "options": ["是", "否"] },
      { "id": "report_annual", "label": "年度财报提供期限 (年后x天)", "tag": "ReportDays_Annual", "type": "number", "value": "45", "formatFn": "chineseNumber" },
      { "id": "report_quarterly", "label": "季度财报提供期限 (季后x天)", "tag": "ReportDays_Quarterly", "type": "number", "value": "30", "formatFn": "chineseNumber" },
      { "id": "report_monthly", "label": "月度财报提供期限 (月后x天)", "tag": "ReportDays_Monthly", "type": "number", "value": "15", "formatFn": "chineseNumber" },
      { "id": "report_budget", "label": "年度预算提供期限 (年后x天)", "tag": "ReportDays_Budget", "type": "number", "value": "45", "formatFn": "chineseNumber" },
      { "id": "hasMFN", "label": "最优惠条款 (MFN)", "tag": "HasMFN", "type": "radio", "options": ["是", "否"] },
      { "id": "hasNewProjectRight", "label": "新项目投资权 (创始人再创业)", "tag": "HasNewProjectRight", "type": "radio", "options": ["是", "否"] }
    ]
  },
  {
    "id": "section_cps",
    "header": { "label": "13. 交割先决条件", "tag": "Section_CPs" },
    "fields": [
      { "id": "cp_warranties", "label": "1. 声明与保证真实准确完整", "tag": "CP_Warranties", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "cp_docs", "label": "2. 签署交易文件(股东协议+新章程)", "tag": "CP_SignDocs", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true, "subFields": [{ "id": "cp_articles_date", "label": "公司章程签订日期", "tag": "CP_ArticlesDate", "type": "date" }, { "id": "cp_sha_date", "label": "股东协议签订日期", "tag": "CP_SHADate", "type": "date" }] },
      { "id": "cp_approval", "label": "3. 股东会批准本次交易", "tag": "CP_Approval", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true, "subFields": [{ "id": "cp_board_size", "label": "董事会总人数", "tag": "CP_BoardSize", "type": "number", "placeholder": "如：5" }, { "id": "cp_founder_directors", "label": "创始股东委派董事数", "tag": "CP_FounderDirectors", "type": "number", "placeholder": "如：2" }] },
      { "id": "cp_aic", "label": "4. 完成工商变更登记", "tag": "CP_AIC", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "cp_key_personnel", "label": "5. 关键人员全职加入", "tag": "CP_KeyPersonnel", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true, "subFields": [{ "id": "cp_labor_term", "label": "劳动合同最低期限(年)", "tag": "CP_LaborTerm", "type": "number", "placeholder": "如：4", "formatFn": "chineseNumber" }] },
      { "id": "cp_no_mac", "label": "6. 无重大不利变化(MAC)", "tag": "CP_NoMAC", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "cp_remittance", "label": "7. 发出汇款通知", "tag": "CP_Remittance", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "cp_closing_notice", "label": "8. 交割条件满足通知", "tag": "CP_ClosingNotice", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "cp_ic_approval", "label": "9. 投资委员会批准", "tag": "CP_ICApproval", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "cp_dd", "label": "10. 尽职调查完成", "tag": "CP_DD", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "cp_founder_holdco", "label": "11. 创始人持股公司承诺函", "tag": "CP_FounderHoldco", "type": "radio", "options": ["适用", "不适用"], "hasParagraphToggle": true },
      { "id": "cp_payment_days", "label": "先决条件满足后付款天数", "tag": "CP_PaymentDays", "type": "number", "value": "10" }
    ]
  },
  {
    "id": "section_veto",
    "header": { "label": "15. 重大事项否决权", "tag": "Section_Veto" },
    "fields": [
      { "id": "veto_subject", "label": "拥有一票否决权的主体", "tag": "VetoSubject", "type": "text", "value": "本轮投资方" },
      { "id": "veto_cap_inc", "label": "1. 增加注册资本/发行新股", "tag": "Veto_IncreaseCapital", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_cap_dec", "label": "2. 减少注册资本/回购股权", "tag": "Veto_DecreaseCapital", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_structure", "label": "3. 修改融资方案/股权结构", "tag": "Veto_Structure", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_rights", "label": "4. 修改股东权利/优先权", "tag": "Veto_AmendRights", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_articles", "label": "5. 修改公司章程", "tag": "Veto_AmendArticles", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_board", "label": "6. 变更董事会人数/产生方式", "tag": "Veto_ChangeBoard", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_senior", "label": "7. 聘用/解聘高管(CEO/CFO等)", "tag": "Veto_SeniorMgmt", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_assets", "label": "8. 重大资产出售/收购/许可", "tag": "Veto_DisposeAssets", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_guarantee", "label": "9. 对外担保/借款", "tag": "Veto_Guarantees", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_related", "label": "10. 关联交易", "tag": "Veto_RelatedTx", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_dividend", "label": "11. 利润分配/分红", "tag": "Veto_Dividends", "type": "radio", "options": ["适用", "不适用"] },
      { "id": "veto_ipo_ma", "label": "12. 上市(IPO)或并购(M&A)方案", "tag": "Veto_IPO_MA", "type": "radio", "options": ["适用", "不适用"] }
    ]
  }
]

;
