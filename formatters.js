/**
 * 格式化函数模块
 * 用于将用户输入的原始数据转换成特定格式
 * 
 * 核心原则：只在"用户输入原始数据，需要程序转换格式"的场景使用格式化函数
 * 大多数情况下用 none，直接替换占位符即可
 */

// ==================== 工具函数 ====================

/**
 * 将数字转换为中文大写
 * @param {number} num - 阿拉伯数字
 * @returns {string} 中文大写数字
 */
function numberToChinese(num) {
  const digits = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
  const units = ['', '拾', '佰', '仟'];
  const bigUnits = ['', '万', '亿', '兆'];
  
  if (num === 0) return '零';
  if (num < 0) return '负' + numberToChinese(-num);
  
  let result = '';
  let unitIndex = 0;
  
  while (num > 0) {
    const section = num % 10000;
    if (section > 0) {
      let sectionStr = '';
      let sectionNum = section;
      let localUnitIndex = 0;
      let hasZero = false;
      
      while (sectionNum > 0) {
        const digit = sectionNum % 10;
        if (digit === 0) {
          if (!hasZero && sectionStr !== '') {
            sectionStr = digits[0] + sectionStr;
          }
          hasZero = true;
        } else {
          sectionStr = digits[digit] + units[localUnitIndex] + sectionStr;
          hasZero = false;
        }
        sectionNum = Math.floor(sectionNum / 10);
        localUnitIndex++;
      }
      
      result = sectionStr + bigUnits[unitIndex] + result;
    }
    num = Math.floor(num / 10000);
    unitIndex++;
  }
  
  // 清理规则
  result = result.replace(/零+/g, '零'); // 多个零合并
  result = result.replace(/零([万亿兆])/g, '$1'); // 零万 -> 万
  result = result.replace(/零+$/, ''); // 结尾的零
  result = result.replace(/^壹拾/, '拾'); // 10-19 的特殊处理
  
  return result || '零';
}

/**
 * 将数字转换为小写中文数字（用于百分比等场景）
 * @param {number} num - 阿拉伯数字
 * @returns {string} 中文数字
 */
function toChineseNumber(num) {
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
  
  if (num >= 0 && num <= 10) {
    return digits[num];
  }
  
  if (num < 20) {
    return '十' + (num % 10 === 0 ? '' : digits[num % 10]);
  }
  
  if (num < 100) {
    const tens = Math.floor(num / 10);
    const ones = num % 10;
    return digits[tens] + '十' + (ones === 0 ? '' : digits[ones]);
  }
  
  // 更大的数字用大写逻辑
  return numberToChinese(num);
}

/**
 * 格式化日期
 * @param {string|Date} dateValue - 日期对象或日期字符串
 * @param {string} format - 格式模板
 * @returns {string} 格式化后的日期
 */
function formatDate(dateValue, format) {
  let date;
  
  if (typeof dateValue === 'string') {
    date = new Date(dateValue);
  } else if (dateValue instanceof Date) {
    date = dateValue;
  } else {
    return '';
  }
  
  if (isNaN(date.getTime())) {
    return '';
  }
  
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  
  return format
    .replace('YYYY', year)
    .replace('YY', String(year).slice(-2))
    .replace('MM', String(month).padStart(2, '0'))
    .replace('M', month)
    .replace('DD', String(day).padStart(2, '0'))
    .replace('D', day);
}

// ==================== 格式化函数 ====================

/**
 * 无格式化，直接返回原值
 */
function formatNone(value) {
  return value;
}

/**
 * 格式化日期为 YYYY年MM月DD日
 * @param {string|Date} dateValue - 日期
 * @returns {string} 2024年01月15日
 */
function formatDateUnderline(dateValue) {
  return formatDate(dateValue, 'YYYY年MM月DD日');
}

/**
 * 格式化日期为 YYYY年MM月
 * @param {string|Date} dateValue - 日期
 * @returns {string} 2024年01月
 */
function formatDateYearMonth(dateValue) {
  return formatDate(dateValue, 'YYYY年MM月');
}

/**
 * 格式化为中文大写数字（带括号）
 * @param {number} numValue - 数字
 * @returns {string} 壹佰（100）
 */
function formatChineseNumber(numValue) {
  const num = Number(numValue);
  if (isNaN(num)) return numValue;
  
  const chinese = numberToChinese(num);
  return `${chinese}（${num}）`;
}

/**
 * 格式化为中文大写数字+万元
 * @param {number} numValue - 数字
 * @returns {string} 壹佰（100）万元
 */
function formatChineseNumberWan(numValue) {
  const num = Number(numValue);
  if (isNaN(num)) return numValue;
  
  const chinese = numberToChinese(num);
  return `${chinese}（${num}）万元`;
}

/**
 * 格式化为完整的金额大写
 * @param {number} numValue - 数字
 * @returns {string} 人民币685000元（大写：陆拾捌万伍仟元整）
 */
function formatAmountWithChinese(numValue) {
  const num = Number(numValue);
  if (isNaN(num)) return numValue;
  
  const chinese = numberToChinese(num);
  return `人民币${num}元（大写：${chinese}元整）`;
}

/**
 * 格式化为条款编号
 * @param {number} numValue - 数字
 * @returns {string} 第五条
 */
function formatArticleNumber(numValue) {
  const num = Number(numValue);
  if (isNaN(num)) return numValue;
  
  const chinese = toChineseNumber(num);
  return `第${chinese}条`;
}

/**
 * 格式化为中文百分比
 * @param {number} numValue - 数字
 * @returns {string} 百分之十
 */
function formatPercentageChinese(numValue) {
  const num = Number(numValue);
  if (isNaN(num)) return numValue;
  
  const chinese = toChineseNumber(num);
  return `百分之${chinese}`;
}

// ==================== 格式化函数映射表 ====================

const FORMAT_FUNCTIONS = {
  'none': formatNone,
  'dateUnderline': formatDateUnderline,
  'dateYearMonth': formatDateYearMonth,
  'chineseNumber': formatChineseNumber,
  'chineseNumberWan': formatChineseNumberWan,
  'amountWithChinese': formatAmountWithChinese,
  'articleNumber': formatArticleNumber,
  'percentageChinese': formatPercentageChinese,
};

// ==================== 白名单验证 ====================

/**
 * 验证 formatFn 是否在白名单中
 * @param {string} formatFn - 格式化函数名称
 * @returns {boolean} 是否有效
 */
function isValidFormatFn(formatFn) {
  return formatFn in FORMAT_FUNCTIONS;
}

/**
 * 获取所有有效的 formatFn 列表
 * @returns {string[]} formatFn 列表
 */
function getValidFormatFns() {
  return Object.keys(FORMAT_FUNCTIONS);
}

/**
 * 应用格式化函数
 * @param {any} value - 输入值
 * @param {string} formatFn - 格式化函数名称
 * @param {object} params - 额外参数（可选）
 * @returns {any} 格式化后的值
 */
function applyFormat(value, formatFn = 'none', params = {}) {
  if (!isValidFormatFn(formatFn)) {
    console.warn(`[Formatters] 未知格式函数: ${formatFn}，使用原值。`);
    console.warn(`[Formatters] 有效的格式函数: ${getValidFormatFns().join(', ')}`);
    return value;
  }
  
  try {
    return FORMAT_FUNCTIONS[formatFn](value, params);
  } catch (error) {
    console.error(`[Formatters] 格式化失败: ${formatFn}`, error);
    return value;
  }
}

// ==================== 导出 ====================

// 兼容浏览器环境
if (typeof window !== 'undefined') {
  window.Formatters = {
    applyFormat,
    isValidFormatFn,
    getValidFormatFns,
    FORMAT_FUNCTIONS,
    // 导出工具函数供测试
    numberToChinese,
    toChineseNumber,
    formatDate,
  };
}

// 兼容 Node.js 环境
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    applyFormat,
    isValidFormatFn,
    getValidFormatFns,
    FORMAT_FUNCTIONS,
    numberToChinese,
    toChineseNumber,
    formatDate,
  };
}
