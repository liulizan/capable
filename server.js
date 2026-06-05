/**
 * Capable（可培）— 后端服务（离线版）
 *
 * 职责：
 * 1. 内置 Capable 五步流程的智能模板引擎
 * 2. 根据用户输入内容的关键词，动态生成能力卡片 + 行动脚本 + 角色扮演邀请
 * 3. 支持 Word (.docx) 和 Excel (.xlsx) 文件上传解析
 * 4. 零外部 API 依赖，无需任何 Key
 */

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const https = require('https');
const http = require('http');
const { URL } = require('url');

const app = express();
const PORT = process.env.PORT || 3000;

// ==================== 中间件 ====================

app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// 文件上传：存入内存（不落盘），限制 10MB，只允许 Word/Excel/Markdown
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 4 * 1024 * 1024 }, // 4MB（Vercel Hobby 4.5MB 上限）
  fileFilter: (_req, file, cb) => {
    const allowed = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',       // .xlsx
      'text/markdown',      // .md
      'text/plain',          // .md 在某些系统上报为 text/plain
    ];
    // macOS / Windows 上 MIME 可能变体，额外放行扩展名
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(file.mimetype) || ext === '.docx' || ext === '.xlsx' || ext === '.md') {
      cb(null, true);
    } else {
      cb(new Error('仅支持 .docx（Word）、.xlsx（Excel）和 .md（Markdown）格式'));
    }
  },
});

// ==================== 会话存储 ====================

// 服务端存储分析上下文，避免泄露到客户端
const sessions = new Map();
const SESSION_TTL = 60 * 60 * 1000; // 1 小时后自动清理

function createSession(analysis, userProfile) {
  const id = `sess_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  sessions.set(id, { analysis, userProfile, createdAt: Date.now() });
  // 定期清理过期会话
  if (sessions.size % 10 === 0) cleanExpiredSessions();
  return id;
}

function getSession(id) {
  const session = sessions.get(id);
  if (!session) return null;
  if (Date.now() - session.createdAt > SESSION_TTL) {
    sessions.delete(id);
    return null;
  }
  return session;
}

function cleanExpiredSessions() {
  const now = Date.now();
  for (const [id, s] of sessions) {
    if (now - s.createdAt > SESSION_TTL) sessions.delete(id);
  }
}

// ==================== 文件解析工具 ====================

/** 解析 .docx → 纯文本 */
async function parseDocx(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  return result.value.trim();
}

/** 解析 .xlsx → 纯文本（遍历所有 Sheet，单元格用 Tab 分隔） */
function parseXlsx(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const parts = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    // 先用 CSV 格式获取原始文本
    const csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false });
    if (csv.trim()) {
      parts.push(`【Sheet: ${sheetName}】\n${csv}`);
    }
  }
  return parts.join('\n\n').trim();
}

// ==================== 内容深度分析引擎 ====================

// 中文停用词（常见虚词，不参与关键词提取）
const STOP_WORDS = new Set([
  '的', '了', '在', '是', '我', '有', '和', '就', '不', '人', '都', '一', '一个',
  '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有', '看', '好',
  '自己', '这', '他', '她', '它', '们', '那', '些', '所', '以', '之', '与', '及',
  '为', '等', '从', '被', '把', '对', '向', '让', '用', '能', '可以', '需要', '应该',
  '但', '而', '且', '或', '如果', '因为', '所以', '因此', '虽然', '然而', '不过',
  '这个', '那个', '什么', '怎么', '怎样', '如何', '为什么', '哪', '吗', '呢', '吧',
  '啊', '哦', '嗯', '哈', '呀', '哇', '啦', '么', '嘛', '呗',
  '我们', '他们', '你们', '她们', '它们', '大家', '别人', '自己',
  '还有', '已经', '比较', '非常', '真的', '特别', '更加', '最', '更',
  '现在', '以后', '以前', '时候', '时间', '今天', '明天', '昨天',
  '通过', '进行', '使用', '利用', '以及', '包括', '其他', '其中',
  '一种', '各种', '不同', '主要', '重要', '相关', '基本', '一般',
  '可能', '可以', '能够', '需要', '必须', '应该', '一定',
  // 英文
  'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
  'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could',
  'should', 'may', 'might', 'can', 'shall', 'to', 'of', 'in', 'for',
  'on', 'with', 'at', 'by', 'from', 'as', 'into', 'through', 'during',
  'before', 'after', 'above', 'below', 'between', 'and', 'but', 'or',
  'nor', 'not', 'so', 'yet', 'both', 'either', 'neither', 'each', 'every',
  'all', 'any', 'few', 'more', 'most', 'other', 'some', 'such', 'no',
  'only', 'own', 'same', 'than', 'too', 'very', 'just', 'because',
  'about', 'also', 'this', 'that', 'these', 'those', 'it', 'its',
]);

/**
 * 内容分析结果：
 * {
 *   title: string,        // 文档标题
 *   domain: string,       // 领域分类
 *   coreConcepts: [],     // 核心概念（高频实词）
 *   keySentences: [],     // 重要句子（含核心概念的句子）
 *   claims: [],           // 文档中的核心观点/论断
 *   methods: [],          // 文档中提出的方法/步骤
 *   docType: string,      // 文档类型
 *   contextHints: {},     // 场景线索
 *   paragraphCount: number,
 *   sentenceCount: number,
 * }
 */
/**
 * 从用户的学习目标中提取关键词
 * 例如 "产品经理，想提升跨部门沟通能力" → ["产品经理","跨部门","沟通"]
 */
function extractGoalTerms(userProfile) {
  if (!userProfile) return [];
  // 先切出"目标"部分：如果有"想/提升/学习"等词，取后面部分
  let goalPart = userProfile;
  const goalSplit = userProfile.split(/[，,、\s]+(?=想|提升|学习|学会|掌握|了解|提高|增强|改善)/);
  if (goalSplit.length > 1) {
    goalPart = goalSplit.slice(1).join('');
  }
  // 剥离意愿词和身份词
  const cleaned = goalPart
    .replace(/想|提升|学会|掌握|了解|学习|提高|增强|加强|改善|改进/g, '')
    .trim();

  if (!cleaned) {
    // fallback: 如果清洗后为空，用原始输入的所有非身份词
    return userProfile
      .replace(/[，,、]/g, ' ')
      .replace(/想|提升|学会|掌握|了解|学习/g, '')
      .trim()
      .split(/\s+/)
      .filter(p => p.length >= 2);
  }

  // 2-4字滑动窗口（仅从目标部分提取）
  const chars = [...cleaned.replace(/\s/g, '')];
  const terms = [];
  for (let len = 4; len >= 2; len--) {
    for (let i = 0; i <= chars.length - len; i++) {
      const term = chars.slice(i, i + len).join('');
      if (term.length >= 2 && !STOP_WORDS.has(term)) {
        terms.push(term);
      }
    }
  }
  const rawParts = cleaned.split(/[\s]+/).filter(p => p.length >= 2);
  // 去重 + 去掉子串（"跨部门沟" 是 "跨部门沟通" 的子串）
  const allTerms = [...new Set([...rawParts, ...terms])];
  return allTerms.filter(t =>
    !allTerms.some(other => other !== t && other.includes(t))
  );
}

/**
 * 计算一段文本与用户学习目标的「相关性得分」
 * 每个目标词在文本中出现一次 +1 分
 */
function scoreGoalRelevance(text, goalTerms) {
  if (!goalTerms || goalTerms.length === 0) return 0;
  const lower = text.toLowerCase();
  let score = 0;
  for (const gt of goalTerms) {
    // 完整匹配得分更高
    if (lower.includes(gt.toLowerCase())) {
      score += gt.length >= 3 ? 3 : 1;
    }
  }
  return score;
}

function analyzeContent(content, userProfile) {
  const clean = content.replace(/\r/g, '').trim();
  const goalTerms = extractGoalTerms(userProfile);

  // --- 1. 提取句子 & 段落 ---
  const sentences = splitSentences(clean).filter(s => s.trim().length > 4);
  const paragraphs = clean.split(/\n\n+/).filter(p => p.trim());
  const paragraphCount = paragraphs.length;
  const totalSentences = sentences.length;

  if (totalSentences === 0) {
    return {
      title: '无标题', domain: '通用能力', coreConcepts: [], keySentences: [],
      claims: [], methods: [], docType: '短文', goalTerms, goalHighlights: [],
      paragraphCount: 0, sentenceCount: 0, wordCount: 0,
      centrality: [], sentences: [],
    };
  }

  // --- 2. TextRank 句子中心性（核心改进） ---
  const centrality = computeSentenceCentrality(sentences);

  // --- 3. 提取有区分度的概念词（TF-IDF 权重排序） ---
  const distinctiveTerms = getDistinctiveTerms(sentences);

  // --- 4. 提取标题：取中心性最高的前3句中最短且信息量最大的 ---
  const topCentralIndices = centrality
    .map((score, idx) => ({ score, idx }))
    .sort((a, b) => b.score - a.score)
    .slice(0, 3);

  let title = '';
  const lines = clean.split('\n').filter(l => l.trim());
  // 优先：取高中心性句子中最像标题的（短、不含标点分隔）
  for (const { idx } of topCentralIndices) {
    const s = sentences[idx].trim();
    if (s.length >= 4 && s.length <= 40 && !/[，。！？；：、]/.test(s)) {
      title = s;
      break;
    }
  }
  if (!title) {
    // fallback: 首行非空文本
    for (const line of lines) {
      const t = line.trim();
      if (t.length >= 2 && t.length <= 60 && !t.startsWith('-') && !t.startsWith('*') && !t.startsWith('#')) {
        title = t.replace(/^[#\s]+/, '').replace(/[#\s]+$/, '');
        if (title.length >= 2) break;
      }
    }
  }
  if (!title) {
    title = sentences[topCentralIndices[0]?.idx || 0]?.slice(0, 40).replace(/\n/g, ' ').trim() + '…' || '无标题';
  }

  // --- 5. 句子多维度评分（中心性为主，句法特征为辅） ---
  const claimMarkers = [
    '核心', '关键', '重要', '本质', '根本', '基础', '原则',
    '因为', '所以', '因此', '导致', '原因', '结果',
    '需要', '必须', '应该', '建议', '推荐', '最好',
    '不是', '而是', '并非', '恰恰', '真正的', '事实上',
    '总结', '概括', '总之', '一句话', '换言之',
  ];

  const scoredSentences = sentences.map((s, idx) => {
    const textLen = s.length;
    const lower = s.toLowerCase();

    // A. 中心性分数（主力）—— 加权 8 倍
    const centralityScore = centrality[idx] * 8;

    // B. 位置分（辅助）
    const posRatio = totalSentences > 1 ? idx / (totalSentences - 1) : 0;
    let positionScore = 0;
    if (posRatio < 0.12) positionScore = 2;
    else if (posRatio < 0.25) positionScore = 1.5;
    else if (posRatio > 0.88) positionScore = 1.5;
    else if (posRatio > 0.75) positionScore = 0.5;

    // C. 定义/判断句式
    const hasDefinition = /是|即|指的是|所谓|就是|是指/.test(s) ? 1.5 : 0;

    // D. 对比句式
    const hasContrast = /不是.*而是|并非.*而是|与其.*不如/.test(s) ? 2 : 0;

    // E. 长度适中
    let lengthScore = 0;
    if (textLen >= 20 && textLen <= 120) lengthScore = 1;

    // F. 目标相关性
    const goalScore = scoreGoalRelevance(s, goalTerms);

    // G. 是否是论断/方法句
    const isClaim = claimMarkers.some(m => s.includes(m)) && textLen > 8;
    const isMethod = /^[\d一二三四五六七八九十]+[\.、．)）]/.test(s.trim()) ||
                     /^第[一二三四五六七八九十\d]+[步条]/.test(s.trim()) ||
                     /步骤[一二三四五六七八九十\d]/.test(s);

    const hasGoal = goalTerms.length > 0;
    const combinedScore = centralityScore + positionScore + hasDefinition +
                          hasContrast + lengthScore +
                          (hasGoal ? goalScore * 2 : 0);

    return { text: s, centralityScore, goalScore, combinedScore, isClaim, isMethod, idx };
  });

  // --- 6. 关键句：按综合分（中心性主导）排序 ---
  const keySentences = scoredSentences
    .sort((a, b) => b.combinedScore - a.combinedScore)
    .slice(0, 10)
    .map(s => s.text);

  // --- 7. 核心观点 ---
  let claims = scoredSentences
    .filter(s => s.isClaim)
    .sort((a, b) => b.combinedScore - a.combinedScore)
    .slice(0, 8)
    .map(s => s.text);

  // --- 8. 方法/步骤 ---
  let methods = scoredSentences
    .filter(s => s.isMethod)
    .sort((a, b) => b.combinedScore - a.combinedScore)
    .slice(0, 8)
    .map(s => s.text);

  // --- 9. 目标高亮句 ---
  const goalHighlights = goalTerms.length > 0
    ? scoredSentences
        .filter(s => s.goalScore >= 2)
        .sort((a, b) => b.goalScore - a.goalScore)
        .slice(0, 5)
        .map(s => s.text)
    : [];

  // --- 10. 核心概念：全文字符级 n-gram 频率 + 区分度排序 ---
  // 直接从清洗后的全文提取所有 2-4 字 n-gram，按频率×长度排序
  const cleanText = clean.replace(/[\s，。！？、；：""''（）\(\)\[\]【】《》〈〉\n\r\t,\.!\?;:"'\(\)\[\]<>|\\\/@#$%^&\*\+=\{\}~`…—\-]+/g, '');
  const conceptFreq = new Map();
  const conceptChars = [...cleanText];

  for (let len = 4; len >= 2; len--) {
    for (let i = 0; i <= conceptChars.length - len; i++) {
      const ngram = conceptChars.slice(i, i + len).join('');
      if (ngram.length < 2) continue;
      // 过滤含过多虚词字符的 n-gram
      const funcCount = [...ngram].filter(c => '的了着过吗呢吧啊呀哇之乎者也所其但而虽且或'.includes(c)).length;
      if (funcCount >= len * 0.5) continue; // 超过一半是虚词 → 跳过
      if (/^[\d\.\,\-\+％％]+$/.test(ngram)) continue;
      conceptFreq.set(ngram, (conceptFreq.get(ngram) || 0) + 1);
    }
  }

  // 按"频率 × 长度"排序（长词+高频 = 更像真实概念）
  // 同时检查与句子的相关性：在越多句子中出现 = 越重要
  const conceptScores = [...conceptFreq.entries()]
    .filter(([, count]) => count >= 2)
    .filter(([term]) => {
      // 过滤无效概念：纯数字、含量词/数词前缀、纯虚词
      if (/^[\d\.\,\-\+％％]+$/.test(term)) return false;
      // 过滤以量词/数词开头的碎片（"个番茄钟"、"4个番茄"）
      if (/^[个只条张本块片件次回趟天年月日时分秒\d一二三四五六七八九十百千万亿]/.test(term)) return false;
      // 过滤纯否定词、纯助词、结构词
      if (/^(不是|就是|而是|也是|还是|都是|只是|没有|这个|那个|什么|怎么|一个|一种|部分|关于|第一|第二|第三|首先|其次|最后|另外|此外|以下|以上|例如|比如)$/.test(term)) return false;
      return true;
    })
    .map(([term, count]) => {
      // 计算句子覆盖率
      let sentCount = 0;
      for (const s of sentences) {
        if (s.includes(term)) sentCount++;
      }
      // 分数 = 频率 × 长度 × 句子覆盖率
      const score = count * term.length * (1 + Math.log(1 + sentCount));
      return { term, count, sentCount, score };
    })
    .sort((a, b) => b.score - a.score);

  let coreConcepts = conceptScores.slice(0, 12).map(c => c.term);
  if (goalTerms.length > 0) {
    const goalRelated = coreConcepts.filter(c =>
      goalTerms.some(g => c.includes(g) || g.includes(c))
    );
    const others = coreConcepts.filter(c => !goalRelated.includes(c));
    coreConcepts = [...goalRelated, ...others].slice(0, 8);
  } else {
    coreConcepts = coreConcepts.slice(0, 8);
  }
  // 去子串：如果"高效沟通"和"效沟通"同时出现，去掉短的子串
  coreConcepts = coreConcepts.filter((c, i, arr) =>
    !arr.some((other, j) => j < i && other.includes(c) && other !== c)
  );

  // --- 11. 领域（用全文高频概念词分类，比 TF-IDF 更准确） ---
  const domainClassifyTerms = conceptScores.slice(0, 25).map(c => c.term);
  const domain = classifyDomain(domainClassifyTerms, clean);

  // --- 12. 文档类型 ---
  let docType = '论述文';
  if (methods.length >= 3) docType = '教程/方法论';
  else if (clean.includes('？') && totalSentences < 20) docType = '问答/短文';
  else if (lines.length > 40 && paragraphCount > 5) docType = '长文/报告';
  else if (/^(#{1,3}\s|第[一二三\d]+[章节篇])/.test(clean)) docType = '结构化文章';

  return {
    title,
    domain,
    coreConcepts,
    keySentences,
    claims,
    methods,
    docType,
    goalTerms,
    goalHighlights,
    paragraphCount,
    sentenceCount: totalSentences,
    wordCount: clean.length,
    centrality,
    sentences,
  };
}

/**
 * 分词：从文本中提取有意义的短语（词/概念）
 *
 * 策略（无外部 NLP 依赖）：
 * - 中文：按标点分段后提取 2-4 字 n-gram，过滤纯虚词碎片
 * - 英文：空白分词，保留 3 字母以上，过滤常见停用词
 * - 用 IDF 思想标记：后续在 analyzeContent 中通过 TF-IDF 权重过滤 "出现在几乎所有句子中的高频泛词"
 */
function tokenize(text) {
  const tokens = [];
  // 按标点/空白切分段落
  const segments = text.split(/[\s，。！？、；：""''（）\(\)\[\]【】《》〈〉\n\r\t,\.!\?;:"'\(\)\[\]<>|\\\/@#$%^&\*\+=\{\}~`…—\-]+/);

  // 纯虚词/助词字符（这些几乎不构成独立概念）
  const functionChars = /[的了着过吗呢吧啊呀哇啦嘛呗么之乎者也所其]/;

  for (const seg of segments) {
    if (!seg || seg.length < 2) continue;

    // 纯英文/数字
    if (/^[a-zA-Z0-9_]+$/.test(seg)) {
      if (seg.length >= 3 && !STOP_WORDS.has(seg.toLowerCase())) {
        tokens.push(seg.toLowerCase());
      }
      continue;
    }

    // 中英混合：拆分英文和中文部分
    const parts = seg.split(/([a-zA-Z0-9_]+)/);
    for (const part of parts) {
      if (!part || part.length < 2) continue;
      if (/^[a-zA-Z0-9_]+$/.test(part)) {
        if (part.length >= 3 && !STOP_WORDS.has(part.toLowerCase())) {
          tokens.push(part.toLowerCase());
        }
      } else {
        // 中文：按长度策略提取
        const chars = [...part];

        if (chars.length <= 4) {
          // 极短片段（≤4字）：直接作为完整概念
          if (chars.length >= 2 && !/^[\d\.\,\-\+％％]+$/.test(part)) {
            tokens.push(part);
          }
        } else if (chars.length <= 8) {
          // 中等片段（5-8字）：保留完整片段 + 提取2-3字子词
          if (!/^[\d\.\,\-\+％％]+$/.test(part)) {
            tokens.push(part); // 完整片段
          }
          // 同时提取 2-3 字 n-gram（覆盖"高效沟通"这样的子概念）
          for (let len = 3; len >= 2; len--) {
            for (let i = 0; i <= chars.length - len; i++) {
              const chunk = chars.slice(i, i + len).join('');
              if (chunk.length >= 2 && !functionChars.test(chunk) && !/^[\d\.\,\-\+％％]+$/.test(chunk)) {
                tokens.push(chunk);
              }
            }
          }
        } else {
          // 长片段（>8字）：只提取 2-3 字 n-gram
          for (let len = 3; len >= 2; len--) {
            for (let i = 0; i <= chars.length - len; i++) {
              const chunk = chars.slice(i, i + len).join('');
              if (chunk.length >= 2 && !functionChars.test(chunk) && !/^[\d\.\,\-\+％％]+$/.test(chunk)) {
                tokens.push(chunk);
              }
            }
          }
        }
      }
    }
  }
  return tokens;
}

/** 分句（中英文通用） */
function splitSentences(text) {
  // 按句末标点切割，保留分隔符
  return text
    .split(/(?<=[。！？\.\!\?])\s*/)
    .flatMap(s => s.split(/\n+/))
    .filter(s => s.trim());
}

// ==================== TF-IDF & TextRank 算法 ====================

/**
 * 为句子集合构建 TF-IDF 词向量
 *
 * 返回：
 * - sentenceVectors: 每个句子的稀疏向量 { term: tfidfWeight }
 * - idfScores: 每个词的 IDF 值
 * - globalVocab: 全局词表（按 IDF 排序）
 */
function computeTFIDF(sentences, tokenizeFn) {
  const N = sentences.length;
  const termDocFreq = new Map(); // 词 → 出现该词的句子数

  // 第一遍：统计 DF（Document Frequency）
  const sentenceTokens = sentences.map(s => {
    const tokens = tokenizeFn(s);
    const unique = new Set(tokens);
    for (const t of unique) {
      termDocFreq.set(t, (termDocFreq.get(t) || 0) + 1);
    }
    return tokens;
  });

  // 计算 IDF: log(N / df)
  const idfScores = new Map();
  for (const [term, df] of termDocFreq) {
    // 出现在 >80% 句子中的词几乎没有区分度
    if (df > N * 0.8) {
      idfScores.set(term, 0);
    } else {
      idfScores.set(term, Math.log((N + 1) / (df + 1)));
    }
  }

  // 第二遍：构建 TF-IDF 向量
  const sentenceVectors = sentenceTokens.map(tokens => {
    const tf = new Map();
    for (const t of tokens) {
      tf.set(t, (tf.get(t) || 0) + 1);
    }
    const vec = {};
    const maxTF = Math.max(...tf.values());
    for (const [t, count] of tf) {
      const idf = idfScores.get(t) || 0;
      if (idf > 0) {
        vec[t] = (count / maxTF) * idf; // 归一化 TF
      }
    }
    return vec;
  });

  // 全局词表：按 IDF 降序排列（IDF越高 = 越有区分度）
  const globalVocab = [...idfScores.entries()]
    .filter(([, idf]) => idf > 0)
    .sort((a, b) => b[1] - a[1])
    .map(([term]) => term);

  return { sentenceVectors, idfScores, globalVocab };
}

/**
 * 计算两个稀疏向量的余弦相似度
 */
function cosineSimilarity(vecA, vecB) {
  let dotProduct = 0;
  let normA = 0;
  let normB = 0;

  for (const key of Object.keys(vecA)) {
    normA += vecA[key] * vecA[key];
    if (vecB[key] !== undefined) {
      dotProduct += vecA[key] * vecB[key];
    }
  }
  for (const val of Object.values(vecB)) {
    normB += val * val;
  }

  if (normA === 0 || normB === 0) return 0;
  return dotProduct / (Math.sqrt(normA) * Math.sqrt(normB));
}

/**
 * TextRank 算法：计算句子在图中的中心性
 *
 * 原理：
 * 1. 句子 = 节点，相似度 > 阈值 = 边
 * 2. PageRank 迭代：每个句子的重要性 = 指向它的句子重要性的加权和
 * 3. 收敛后，中心性最高的句子 = 全文最核心的句子
 *
 * @param {string[]} sentences - 句子列表
 * @param {number} simThreshold - 相似度阈值（低于此值不连边）
 * @param {number} damping - PageRank 阻尼系数
 * @returns {number[]} 每个句子的中心性分数（已归一化到 0-1）
 */
function computeSentenceCentrality(sentences, simThreshold = 0.08, damping = 0.85) {
  const N = sentences.length;
  if (N <= 1) return N === 1 ? [1] : [];

  // 1. TF-IDF 向量化
  const { sentenceVectors } = computeTFIDF(sentences, tokenize);

  // 2. 构建相似度矩阵（只保留超过阈值的边）
  const simMatrix = Array.from({ length: N }, () => new Array(N).fill(0));
  const outSum = new Array(N).fill(0); // 每个节点的出边权重和

  for (let i = 0; i < N; i++) {
    for (let j = i + 1; j < N; j++) {
      const sim = cosineSimilarity(sentenceVectors[i], sentenceVectors[j]);
      if (sim > simThreshold) {
        simMatrix[i][j] = sim;
        simMatrix[j][i] = sim;
        outSum[i] += sim;
        outSum[j] += sim;
      }
    }
  }

  // 3. PageRank 迭代
  let scores = new Array(N).fill(1 / N);
  const baseScore = (1 - damping) / N;

  for (let iter = 0; iter < 50; iter++) {
    const newScores = new Array(N).fill(baseScore);
    let maxDelta = 0;

    for (let i = 0; i < N; i++) {
      let incomingSum = 0;
      for (let j = 0; j < N; j++) {
        if (i !== j && outSum[j] > 0) {
          incomingSum += (simMatrix[j][i] / outSum[j]) * scores[j];
        }
      }
      newScores[i] += damping * incomingSum;
      maxDelta = Math.max(maxDelta, Math.abs(newScores[i] - scores[i]));
    }

    scores = newScores;
    if (maxDelta < 1e-6) break; // 收敛
  }

  // 4. 归一化到 0-1
  const maxScore = Math.max(...scores);
  if (maxScore > 0) {
    scores = scores.map(s => s / maxScore);
  }

  return scores;
}

/**
 * 获取全局 TF-IDF 词表（用于概念提取）
 * 返回按区分度排序的词列表
 */
function getDistinctiveTerms(sentences) {
  const { globalVocab, idfScores } = computeTFIDF(sentences, tokenize);
  // 过滤：只保留长度适中（2-8字）且 IDF > 0 的词
  return globalVocab.filter(t => {
    if (t.length < 2 || t.length > 8) return false;
    // 纯数字/标点
    if (/^[\d\.\,\-\+％％]+$/.test(t)) return false;
    return idfScores.get(t) > 0.5; // IDF 太低 = 太泛
  });
}

/** 根据高频词分类领域 */
function classifyDomain(sortedTerms, content) {
  const domainRules = [
    { domain: '沟通与表达',    words: ['沟通', '表达', '汇报', '演讲', '谈判', '说服', '会议', '对话', '交流', '听众', '发言', '讲话', '提问', '倾听', '反馈', '冲突'] },
    { domain: '时间与效率',    words: ['时间', '效率', '拖延', '计划', '番茄', '优先级', 'deadline', '日程', '任务', '专注', '精力', '习惯', '打卡', '自律'] },
    { domain: '学习与认知',    words: ['学习', '记忆', '费曼', '复习', '读书', '知识', '理解', '练习', '刻意', '反馈', '能力', '技能', '认知', '思维', '方法', '技巧'] },
    { domain: '写作与输出',    words: ['写作', '文案', '文章', '笔记', '总结', '报告', '文档', '邮件', '逻辑', '结构', '标题', '开头', '结尾', '段落', '修改'] },
    { domain: '领导与管理',    words: ['领导', '管理', '团队', '授权', '目标', 'OKR', 'KPI', '激励', '指导', '决策', '战略', '执行', '绩效', '招聘', '培养'] },
    { domain: '产品与设计',    words: ['产品', '用户', '设计', '需求', '体验', '交互', '原型', '迭代', '功能', '痛点', '场景', '数据', '增长', '留存', '转化'] },
    { domain: '技术与编程',    words: ['代码', '编程', '算法', '架构', '系统', '服务', '接口', '数据库', '前端', '后端', '测试', '部署', '性能', '安全', '开源'] },
    { domain: '营销与增长',    words: ['营销', '增长', '用户', '流量', '转化', '品牌', '内容', '社交', '广告', 'SEO', '文案', '粉丝', '传播', '裂变', '渠道'] },
    { domain: '心理与情商',    words: ['情绪', '心理', '焦虑', '压力', '幸福', '关系', '边界', '共情', '自我', '成长', '心态', '正念', '冥想', '自信'] },
    { domain: '数据与分析',    words: ['数据', '分析', '指标', '报表', '图表', '趋势', '模型', '预测', '统计', 'Excel', 'SQL', 'BI', '可视化', '维度'] },
    { domain: '职场与成长',    words: ['职场', '职业', '面试', '简历', '跳槽', '晋升', '薪资', '同事', '老板', '公司', '行业', '转型', '规划'] },
  ];

  const scores = {};
  for (const rule of domainRules) {
    let score = 0;
    const lowerContent = content.toLowerCase();
    for (const w of rule.words) {
      // 在 TF-IDF 排序词中匹配，位置越靠前权重越高
      const exactIdx = sortedTerms.findIndex(t => t === w);
      const partialIdx = sortedTerms.findIndex(t => t.includes(w) && t !== w);
      if (exactIdx >= 0) score += Math.max(10 - exactIdx, 3); // 排名第1=10分，排名第10=3分
      else if (partialIdx >= 0) score += Math.max(5 - Math.floor(partialIdx / 2), 1);

      // 在全文中匹配（封顶 2，辅助判断）
      const count = (lowerContent.match(new RegExp(w, 'g')) || []).length;
      score += Math.min(count, 2);
    }
    if (score > 0) scores[rule.domain] = score;
  }

  // 取最高分；同分时偏好第一个匹配的领域（规则按优先级排列）
  const best = Object.entries(scores).sort((a, b) => b[1] - a[1])[0];
  return best ? best[0] : '通用能力';
}
/** 从分析结果生成「角色扮演场景」——有目标时与目标对齐 */
function inferScenario(analysis, userProfile) {
  const { domain, coreConcepts, goalTerms, goalHighlights, title } = analysis;
  const hasGoal = goalTerms && goalTerms.length > 0;
  const role = userProfile ? userProfile.split(/[,，、]/)[0] : '你';

  // 有目标时：用目标构建更具体的场景
  if (hasGoal && goalHighlights.length > 0) {
    const goalStr = goalTerms.filter(t => t.length >= 3).slice(0, 2).join('、');
    const highlight = goalHighlights[0].replace(/^[\d\.\s、①②③④⑤]+/, '').slice(0, 60);
    return {
      scenario: `你刚学完「${title}」中和「${goalStr}」最相关的内容（核心观点：${highlight}）。现在你的同事向你请教，问"学了这些对我有什么实际帮助"，请你用自己的话，结合你的真实工作场景给出回答`,
      partner: '你的同事',
      role,
    };
  }

  // 基于领域和内容定制场景
  const scenarios = {
    '沟通与表达': `你刚学完「${title}」的内容，明天要在部门周会上向5位同事做3分钟的工作进展汇报。你需要在短时间内抓住重点，让每个人都清楚你的进度和遇到的困难`,
    '时间与效率': `你的同事注意到你最近效率变高了，问你秘诀是什么。你需要用「${title}」中提到的1-2个方法，用通俗的语言解释给他听`,
    '学习与认知': `你的朋友说"我看了很多书但什么也记不住"，根据「${title}」的内容，你该如何用你自己的话帮他解决这个问题`,
    '写作与输出': `领导让你写一份本周工作总结邮件，抄送部门所有人。你刚学完「${title}」，想用里面的方法把邮件写得清晰有力`,
    '领导与管理': `你的团队成员在项目中犯了一个错误，你需要在1对1谈话中既指出问题又不打击他的积极性——运用「${title}」的核心思路`,
    '产品与设计': `需求评审会上，研发对你的方案提出了3个技术实现上的质疑。你需要用「${title}」中的方法有效回应`,
    '技术与编程': `Code Review 时，同事对你的一段代码提出了优化建议，但你不同意。你需要用「${title}」中学到的沟通技巧表达你的观点`,
    '营销与增长': `月底复盘，老板问你"为什么这个月的转化率下降了"，你需要基于「${title}」的分析框架给出回答`,
    '心理与情商': `朋友向你倾诉工作压力很大，快要崩溃了。你刚学过「${title}」，想用里面的方法帮他梳理情绪`,
    '数据与分析': `你用「${title}」的方法做了一份数据分析报告，现在要在全员会议上用3分钟讲清楚核心结论`,
    '职场与成长': `年度绩效面谈时，领导问你"今年最大的成长是什么"，你需要用「${title}」中的框架来组织你的回答`,
  };

  const scenario = scenarios[domain] ||
    `你刚学完「${title}」的核心内容。你的朋友/同事对此很好奇，问你"学到了什么，有用吗"。请用你自己的话，结合${role}的实际工作/学习场景来回答`;

  const partnerMap = {
    '沟通与表达': '部门同事',
    '时间与效率': '你的同事',
    '学习与认知': '你的朋友',
    '写作与输出': '你的领导',
    '领导与管理': '你的团队成员',
    '产品与设计': '研发工程师',
    '技术与编程': '你的同事',
    '营销与增长': '你的老板',
    '心理与情商': '你的朋友',
    '数据与分析': '你的老板',
    '职场与成长': '你的领导',
  };

  const partner = partnerMap[domain] || '对方';

  return { scenario, partner, role };
}

// ==================== 生成函数 ====================

/**
 * 提炼中心论点：优先取"定义/判断句 + 位置靠前 + 含强标记"的句子
 */
function inferOneLiner(analysis) {
  const { keySentences, goalTerms, goalHighlights, centrality, sentences, coreConcepts, title } = analysis;
  const hasGoal = goalTerms && goalTerms.length > 0;

  // --- 从中心性最高的句子中找核心论点 ---
  // 取中心性排名前 5 的句子
  const topByCentrality = (centrality && sentences)
    ? centrality
        .map((score, idx) => ({ score, text: sentences[idx], idx }))
        .sort((a, b) => b.score - a.score)
        .slice(0, 5)
        .map(s => s.text)
    : keySentences.slice(0, 5);

  // --- 有目标：目标桥接（优先用目标高亮句） ---
  if (hasGoal && goalHighlights.length > 0) {
    const best = goalHighlights[0].replace(/^[\d\.\s、①②③④⑤\*\-\#]+/, '').slice(0, 90);
    const goalStr = goalTerms.filter(t => t.length >= 3).slice(0, 2).join('、');
    return `为达成「${goalStr}」的目标，本文最关键的启示是：${best}`;
  }
  if (hasGoal && topByCentrality.length > 0) {
    const best = topByCentrality[0].replace(/^[\d\.\s、①②③④⑤\*\-\#]+/, '').slice(0, 90);
    const goalStr = goalTerms.filter(t => t.length >= 3).slice(0, 2).join('、');
    return `为达成「${goalStr}」的目标，本文的核心要点：${best}`;
  }

  // --- 无目标：从高中心性句子中找"定义/判断/对比"句 ---
  // 这类句子通常是作者的核心论点
  const insightPatterns = [
    /(核心|关键|本质|真正|最重要|根本)的?(是|在于)[^，。！？]{5,70}/,
    /(不是.{2,25})而是[^，。！？]{5,70}/,
    /总结[^，。！？]{5,70}/,
    /一句话[^，。！？]{5,70}/,
  ];

  for (const pattern of insightPatterns) {
    for (const s of topByCentrality) {
      const m = s.match(pattern);
      if (m) {
        return s.replace(/^[\d\.\s、①②③④⑤\*\-\#]+/, '').slice(0, 100);
      }
    }
  }

  // 次优：中心性最高的非纯结构句（排除纯列举/过渡句）
  const structurePatterns = /^(首先|其次|最后|另外|此外|接下来|下面|以下是|如下|第[一二三\d]+[步条章])/;
  for (const s of topByCentrality) {
    const cleaned = s.replace(/^[\d\.\s、①②③④⑤\*\-\#]+/, '').trim();
    if (!structurePatterns.test(cleaned) && cleaned.length >= 15) {
      return cleaned.slice(0, 100);
    }
  }

  // 兜底
  if (topByCentrality.length > 0) {
    return topByCentrality[0].slice(0, 100);
  }
  return `掌握「${coreConcepts[0] || title}」的核心要义并付诸实践`;
}

/**
 * 生成能力名称——如果有目标，体现"服务于目标"的导向
 */
function inferAbilityName(analysis) {
  const { domain, coreConcepts, claims, methods, goalTerms } = analysis;
  const hasGoal = goalTerms && goalTerms.length > 0;

  let concept = '';
  // 有目标时：优先取与目标相关的概念
  if (hasGoal) {
    const goalRelatedTerm = coreConcepts.find(c =>
      goalTerms.some(g => c.includes(g) || g.includes(c))
    );
    if (goalRelatedTerm) {
      concept = goalRelatedTerm;
    }
  }
  // fallback
  if (!concept && methods.length > 0) {
    concept = methods[0].trim().replace(/^[\d一二三四五六七八九十]+[\.、．)）\s]+/, '').slice(0, 20);
  }
  if (!concept) {
    // 取长度>=3 且不含纯虚词的概念
    const actionableTerms = coreConcepts.filter(c =>
      c.length >= 3 && c.length <= 8 &&
      !/^[的了着过吗呢吧啊呀之乎者也所其但而虽且或为被把对向让用能可]+$/.test(c)
    );
    if (actionableTerms.length >= 2) concept = actionableTerms[0] + '·' + actionableTerms[1];
    else if (actionableTerms.length === 1) concept = actionableTerms[0];
  }
  if (!concept && claims.length > 0) {
    concept = claims[0].replace(/[，。！？\n]/g, '').slice(0, 18);
  }
  if (!concept) concept = analysis.title.slice(0, 18);

  const prefixMap = {
    '沟通与表达': '学会', '时间与效率': '掌握', '学习与认知': '习得',
    '写作与输出': '提升', '领导与管理': '修炼', '产品与设计': '建立',
    '技术与编程': '精通', '营销与增长': '玩转', '心理与情商': '培养',
    '数据与分析': '驾驭', '职场与成长': '打造',
  };
  const prefix = prefixMap[domain] || '掌握';

  return `${prefix}「${concept}」`;
}

/** 适用边界 */
function inferBoundary(analysis) {
  const { domain } = analysis;
  const m = {
    '沟通与表达': '工作汇报 / 跨部门协调 / 公开演讲 / 1对1深度对话｜不适用：法律/公文等规范文书',
    '时间与效率': '日常工作排期 / 多项目并行 / 个人习惯养成｜不适用：突发危机应急响应',
    '学习与认知': '新领域入门 / 备考复习 / 专业技能精进｜不适用：纯体能型技能训练',
    '写作与输出': '工作邮件 / 周报月报 / 技术文档 / 自媒体文章｜不适用：文学创作',
    '领导与管理': '团队日常管理 / 项目推进 / 绩效面谈｜不适用：跨国跨文化管理',
    '产品与设计': '互联网产品规划 / 功能迭代 / 体验优化｜不适用：硬件/工业设计',
    '技术与编程': '软件开发 / 系统设计 / 代码审查｜不适用：嵌入式底层开发',
    '营销与增长': '线上获客 / 内容营销 / 品牌建设｜不适用：线下渠道管理',
    '心理与情商': '日常人际关系 / 情绪觉察 / 冲突调解｜不适用：临床心理治疗',
    '数据与分析': '业务报表解读 / 数据驱动决策 / A/B测试｜不适用：机器学习建模',
    '职场与成长': '职业规划 / 求职面试 / 技能提升｜不适用：创业/自由职业',
  };
  return m[domain] || `${domain}领域的实际工作/学习场景｜不适用：完全不相关的领域`;
}

/**
 * 核心逻辑——将原文观点提炼为"原则 + 怎么做"的行动指南
 * 每条格式：标题式原则 → 一句话解释如何应用
 */
function inferLogic(analysis) {
  const { claims, methods, centrality, sentences, coreConcepts } = analysis;
  const logic = [];

  // 清理函数
  const cleanText = (s) => s.trim()
    .replace(/^[#\*\-\d一二三四五六七八九十]+[\.、．)）\s]+/, '')
    .replace(/^[①②③④⑤⑥⑦⑧⑨⑩]/, '')
    .replace(/^第[一二三四五六七八九十\d]+[步条章][：:]?\s*/, '')
    .replace(/[，。！？；：\n]+$/, '')
    .slice(0, 80);

  // 用 TF-IDF 向量计算句子间的语义相似度（替代字符级 Jaccard）
  const allCandidates = [...new Set([...claims, ...methods])]; // 去重原始文本
  // 如果有中心性数据，按中心性排序
  if (centrality && sentences) {
    const scoredCandidates = allCandidates
      .map(text => {
        const idx = sentences.indexOf(text);
        return { text, centrality: idx >= 0 ? centrality[idx] : 0 };
      })
      .sort((a, b) => b.centrality - a.centrality);
    allCandidates.length = 0;
    allCandidates.push(...scoredCandidates.map(c => c.text));
  }

  // 构建候选句的 TF-IDF 向量用于去重
  let candidateVectors = [];
  if (allCandidates.length > 0) {
    const { sentenceVectors } = computeTFIDF(allCandidates, tokenize);
    candidateVectors = sentenceVectors;
  }

  // 存储已添加的逻辑条目对应的候选索引，用于去重
  const addedIndices = new Set();

  // 去重：余弦相似度 > 0.65 视为重复
  const addLogic = (text, idx) => {
    const c = cleanText(text);
    if (c.length < 10) return false;
    // 检查是否已经添加过完全相同的文本
    if (logic.includes(c)) return false;

    // 用余弦相似度检查是否与已有逻辑重复
    if (candidateVectors.length > 0 && idx < candidateVectors.length) {
      const vec = candidateVectors[idx];
      for (const addedIdx of addedIndices) {
        if (addedIdx < candidateVectors.length) {
          const sim = cosineSimilarity(vec, candidateVectors[addedIdx]);
          if (sim > 0.65) return false;
        }
      }
    }

    logic.push(c);
    addedIndices.add(idx);
    return true;
  };

  // 优先级：对比句 > 定义句 > 其他
  const contrastCandidates = allCandidates.filter(c => /不是.*而是|并非.*而是|与其.*不如/.test(c));
  for (let i = 0; i < contrastCandidates.length && logic.length < 3; i++) {
    const idx = allCandidates.indexOf(contrastCandidates[i]);
    addLogic(contrastCandidates[i], idx);
  }

  if (logic.length < 3) {
    const defCandidates = allCandidates.filter(c => /是|即|指的是|所谓/.test(c) && !/不是.*而是/.test(c));
    for (let i = 0; i < defCandidates.length && logic.length < 3; i++) {
      const idx = allCandidates.indexOf(defCandidates[i]);
      addLogic(defCandidates[i], idx);
    }
  }

  if (logic.length < 3) {
    for (let i = 0; i < allCandidates.length && logic.length < 3; i++) {
      addLogic(allCandidates[i], i);
    }
  }

  // 兜底
  while (logic.length < 3) {
    const fallbacks = [
      `「${coreConcepts[0] || '核心概念'}」的关键不在于'知道'而在于'做到'——每次遇到相关场景，有意识地用一次`,
      `把「${coreConcepts[1] || '关键方法'}」拆成最小动作单元：第一步只需5分钟就能完成`,
      `用"输入→内化→输出"三步检验学习效果：能用自己的话向别人讲清楚 = 真的懂了`,
    ];
    const fb = fallbacks[logic.length] || fallbacks[0];
    logic.push(fb);
  }

  return logic.slice(0, 3);
}

/**
 * 常见踩坑点——从文档逻辑反向推导
 */
function inferPitfalls(analysis) {
  const { coreConcepts, claims, methods } = analysis;
  const pitfalls = [];

  const c1 = coreConcepts[0] || '核心概念';
  const c2 = coreConcepts[1] || '关键方法';

  // 坑1：基于文档结构推断
  if (methods.length >= 3) {
    pitfalls.push(`❌ 只记住了步骤清单，却没有理解每一步背后的"为什么"——遇到变体场景就不知道该怎么办`);
  } else if (claims.length >= 3) {
    pitfalls.push(`❌ 读的时候觉得"说得太对了"，但合上文档后一个具体的行动都没有——知道 ≠ 做到`);
  } else {
    pitfalls.push(`❌ 把「${c1}」当成万能公式，在任何场景都生搬硬套——没有一种方法适用于所有情况`);
  }

  // 坑2：基于核心概念关系
  if (coreConcepts.length >= 2) {
    pitfalls.push(`❌ 过于关注「${c1}」而忽略了「${c2}」——这两个概念是互补的，单独用效果减半`);
  } else {
    pitfalls.push(`❌ 学完就停——「${c1}」需要在实际场景中至少用3次才能真正内化`);
  }

  return pitfalls;
}

function generateCapabilityCard(analysis, userProfile, cardData) {
  const oneLiner = cardData.oneLiner;
  const abilityName = cardData.abilityName;
  const boundary = inferBoundary(analysis);
  const logic = cardData.logic;
  const pitfalls = cardData.pitfalls;
  const hasGoal = analysis.goalTerms && analysis.goalTerms.length > 0;

  // 用 TF-IDF 筛选真正的关键概念（长度适中、区分度高）
  const keyConcepts = analysis.coreConcepts
    .filter(c => c.length >= 2 && c.length <= 8)
    .filter(c => !/^[\d\.\,\-\+％％]+$/.test(c)) // 排除纯数字
    .slice(0, 4);

  let card = `### 【能力卡片】\n\n`;
  card += `> **💡 一句话核心：** ${oneLiner}\n\n`;

  // 有目标时：展示"目标桥接"——从文档到你目标的路径
  if (hasGoal && analysis.goalHighlights.length > 0) {
    const goalStr = analysis.goalTerms.filter(t => t.length >= 3).slice(0, 2).join('、');
    const highlightCount = analysis.goalHighlights.length;
    card += `**🎯 目标桥接：** 本文中有 **${highlightCount} 处**内容与你的目标「${goalStr}」直接相关，以下提炼围绕你的目标展开。\n\n`;
  }

  card += `【🎯 能力名称】：${abilityName}\n\n`;
  card += `【📍 适用边界】：\n${boundary}\n\n`;
  card += `【🧭 核心行动指南】：`;
  logic.forEach((l, i) => {
    card += `\n${i + 1}. ${l}`;
  });
  card += `\n\n【⚠️ 常见踩坑点】：`;
  pitfalls.forEach((p, i) => {
    card += `\n${i + 1}. ${p}`;
  });
  if (keyConcepts.length > 0) {
    card += `\n\n【🔑 关键概念】：${keyConcepts.join('  ·  ')}`;
  }

  return card;
}

/**
 * 生成行动脚本——基于卡片已提取的 logic/pitfalls/oneLiner
 * 每个部分直接引用卡片的具体结论，不再用独立模板
 *
 * @param {object} card - { logic: string[], pitfalls: string[], oneLiner: string, abilityName: string }
 */
function generateActionScript(analysis, userProfile, card) {
  const { coreConcepts, goalTerms, goalHighlights, title } = analysis;
  const hasGoal = goalTerms && goalTerms.length > 0;
  const goalStr = hasGoal ? goalTerms.filter(t => t.length >= 3).slice(0, 2).join('') : '';
  const role = userProfile ? userProfile.split(/[,，、]/)[0] : '你';

  const logic = card.logic;
  const pitfalls = card.pitfalls;
  const oneLiner = card.oneLiner;
  const c1 = coreConcepts[0] || '核心理念';
  const c2 = coreConcepts[1] || '配套方法';

  // ========== ① 微练习 ==========
  // 直接基于 logic[0] 设计练习
  let microPractice = '';
  const L1 = logic[0] || '';
  const L1short = L1.length > 45 ? L1.slice(0, 42) + '…' : L1;

  if (L1 && L1.length > 10) {
    microPractice = [
      `**依据：** 能力卡片中的核心行动指南 #1——「${L1short}」`,
      ``,
      `**🎯 你的任务：**`,
      `拿出一张纸（或打开备忘录），完成以下 3 步：`,
      `① 用自己的话把上面这条原则改写成一个"如果…就…"指令`,
      `② 写出一个你**明天就会遇到**的具体场景，可以触发这条指令`,
      `③ 写下你预期会发生什么变化`,
      ``,
      `**⏱️ 限时 5 分钟。现在就开始。**`,
    ].join('\n');
  } else {
    microPractice = [
      `**🎯 你的任务：**`,
      `用手机录音 90 秒，向一个完全没读过「${title}」的人解释：`,
      `这篇文章最核心的一个观点是什么？为什么它值得你花时间读？`,
      ``,
      `**规则：** 不许回看原文——只说你真正理解并记住的。`,
      `**⏱️ 限时 90 秒。**`,
    ].join('\n');
  }

  // ========== ② 触发卡 ==========
  // 基于 pitfalls[0] 设计触发器
  let triggerCard = '';
  const P1 = pitfalls[0] || '';
  const pitfallCore = P1.replace(/^❌\s*/, '').replace(/[「」]/g, '').slice(0, 40);

  if (P1 && P1.length > 8) {
    const shortPitfall = P1.replace(/^❌\s*/, '').slice(0, 50) + (P1.length > 50 ? '…' : '');
    const triggerWhen = [
      `当你发现自己又陷入了"${pitfallCore}…"的苗头时`,
      `当你准备把学到的东西"先收藏，以后再看"的时候`,
      `当你发现自己在用旧习惯处理问题，而不是用卡片中的新方法时`,
      `当你觉得"今天太忙了，明天再练"的时候`,
    ];
    const when = triggerWhen[Math.floor(Math.random() * triggerWhen.length)];

    triggerCard = [
      `**⚠️ 警戒信号：** 能力卡片提醒过你——"${shortPitfall}"`,
      ``,
      `**触发条件：** ${when}`,
      ``,
      `**执行动作：**`,
      `1. 心里默念"停"——打断自动导航模式`,
      `2. 快速回忆卡片中最高优先级的那条指南：**「${L1short}」**`,
      `3. 用一个微小的动作切换模式——哪怕只是站起来喝口水、在备忘录打一行字`,
      `4. 睡前花 30 秒记录：今天成功避开了这个坑吗？`,
    ].join('\n');
  } else {
    triggerCard = [
      `**触发条件：** 明天早上打开电脑/开始工作前`,
      ``,
      `**执行动作：**`,
      `花 1 分钟问自己："今天哪个任务可以用上「${c1}」的思路？"`,
      `把它写在便签上，贴在屏幕边——全天可见。`,
    ].join('\n');
  }

  // ========== ③ 自检题 ==========
  // 基于 logic[1] 和 pitfalls[1] 生成，选项随机排列
  const L2 = logic[1] || L1;
  const L2short = L2.length > 45 ? L2.slice(0, 42) + '…' : L2;
  const P2 = pitfalls[1] || P1;

  // 随机排列选项的工具函数
  function shuffleOptions(options) {
    const indices = [0, 1, 2, 3];
    // Fisher-Yates 洗牌
    for (let i = indices.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [indices[i], indices[j]] = [indices[j], indices[i]];
    }
    const letters = ['A', 'B', 'C', 'D'];
    return {
      shuffled: indices.map(i => ({ letter: letters[indices.indexOf(i)], text: options[i] })),
      correctLetter: letters[indices.indexOf(0)], // 正确答案始终是 options[0]
    };
  }

  // 选择题1：检验对核心指南的理解
  const q1Options = [
    '理解其底层逻辑后，立即在实际场景中试用并观察结果',
    '反复阅读相关内容，直到能完整复述',
    '先查找是否有其他权威资料支持这个观点再行动',
    '等条件成熟时再一次性全面应用',
  ];
  const q1 = shuffleOptions(q1Options);
  const q1Answer = q1.correctLetter;
  const q1Explanation = '卡片强调"行动优先、在实践中调整"，理解的目的是为了应用，而非为了记住。停留在认知层面、等待更权威资料、或等条件成熟都是拖延的体面借口。';

  // 选择题2：检验对踩坑点的警觉
  const q2Options = [
    '设置一个外部提醒机制，在关键时刻触发自查',
    '在做之前先写出完美的详细计划',
    '相信自己的直觉，不需要刻意防范',
    '等到这个问题自然消失',
  ];
  const q2 = shuffleOptions(q2Options);
  const q2Answer = q2.correctLetter;
  const q2Explanation = '陷阱之所以是陷阱，正是因为人在其中时往往不自知。外部提醒（便签、闹钟、伙伴监督）比"我相信自己会注意"可靠得多。';

  let script = `---

### 【个性化行动脚本】

**① ⚡ 5分钟微练习**

${microPractice}

**② 🛡️ 防踩坑触发卡**

${triggerCard}

**③ 📝 自检题**

*选择题1（检验你对核心指南的理解）：*
能力卡片提出了一条关键行动指南——「${L2short}」。以下哪个选项**最准确地**体现了这条指南的含义？
${q1.shuffled.map(o => `${o.letter}. ${o.text}`).join('\n')}

> **正确答案：${q1Answer}**
> **解析：** ${q1Explanation}

*选择题2（检验你对踩坑点的警觉）：*
卡片中特别提醒了以下陷阱——「${P2.replace(/^❌\s*/, '')}」。避免这个陷阱最有效的方式是：
${q2.shuffled.map(o => `${o.letter}. ${o.text}`).join('\n')}

> **正确答案：${q2Answer}**
> **解析：** ${q2Explanation}

*简答题：*
能力卡片的"一句话核心"是：「${oneLiner.slice(0, 70)}」

请用你自己的经历来回应：
${hasGoal ? `(1) 结合你的目标「${goalStr}」，这个核心观点对你最有价值的部分是什么？为什么？` : '(1) 这个核心观点让你联想到了自己的哪次经历？那次经历印证了还是反驳了它？'}
(2) 基于以上反思，你明天会采取的**第一个具体行动**是什么？（须包含：什么时候、做什么、和谁有关）`;

  return script;
}

function generateRoleplayInvitation(analysis, userProfile) {
  const { scenario, partner, role } = inferScenario(analysis, userProfile);

  return `---\n\n### 【AI 场景带练】\n\n现在开始角色扮演。我扮演 ${partner}，你扮演 ${role}。情境：${scenario}。请先说你的第一反应：`;
}

function generateRoleplayFeedback(userReply, analysis) {
  const replyLen = userReply.trim().length;
  const coreConcept = analysis ? (analysis.coreConcepts[0] || '核心理念') : '核心理念';
  let goodPoint, missingPoint, suggestion;

  if (replyLen < 15) {
    goodPoint = '你快速给出了回应，没有犹豫——在真实场景中，及时回应比完美回应更重要';
    missingPoint = `回复过于简短，没有体现出你对「${coreConcept}」的理解——对方感觉不到你的思考过程`;
    suggestion = '试试"观点+依据+行动"三段式：一句话亮观点 → 一句话说为什么 → 一句话说接下来做什么';
  } else if (replyLen < 60) {
    goodPoint = '你清晰地表达了立场，方向是对的';
    missingPoint = '缺少一个具体的例子或细节——抽象的观点不容易让对方信服';
    suggestion = `在关键论点后加一个你自己的真实经历——比如"上次我尝试用「${coreConcept}」的思路处理类似问题，结果是…"——立刻让观点从"道理"变成"经验"`;
  } else {
    goodPoint = '你的回复有结构、有内容，能看出你确实消化了文档的核心概念';
    missingPoint = '结尾缺少一个明确的"下一步"承诺——好的沟通以行动收尾，而不是以观点收尾';
    suggestion = `加一句"那我接下来的做法是…"——把「${coreConcept}」从一个抽象概念变成一个明天就能看到结果的动作`;
  }

  return `**✅ 做得好的一点：**
${goodPoint}

**⚠️ 遗漏的关键点：**
${missingPoint}

**💡 一句改进建议：**
${suggestion}`;
}

// ==================== API 接口 ====================

// ==================== 模式 1：总结文章重点 ====================

/**
 * 对句子做简单的层次聚类（基于余弦相似度）
 * 返回每个句子所属的簇编号
 */
function clusterSentencesBySimilarity(sentenceVectors, threshold = 0.15) {
  const N = sentenceVectors.length;
  const clusters = new Array(N).fill(-1);
  let clusterId = 0;

  for (let i = 0; i < N; i++) {
    if (clusters[i] >= 0) continue; // 已分配
    clusters[i] = clusterId;

    for (let j = i + 1; j < N; j++) {
      if (clusters[j] >= 0) continue;
      const sim = cosineSimilarity(sentenceVectors[i], sentenceVectors[j]);
      if (sim > threshold) {
        clusters[j] = clusterId;
      }
    }
    clusterId++;
  }
  return clusters;
}

function generateSummary(analysis, userProfile) {
  const { title, sentences, centrality, coreConcepts } = analysis;
  if (!sentences || sentences.length === 0) {
    return '内容过短，无法生成摘要。请提供更长的文章。';
  }

  // 1. 取中心性 top 50% 的句子
  const topN = Math.max(5, Math.ceil(sentences.length * 0.5));
  const ranked = sentences
    .map((text, idx) => ({ text, idx, centrality: centrality ? centrality[idx] : 0 }))
    .sort((a, b) => b.centrality - a.centrality)
    .slice(0, topN)
    .sort((a, b) => a.idx - b.idx); // 按原文顺序排列

  // 2. 基于结构标记 + 共享关键词的混合聚类
  const keywords = (coreConcepts || []).filter(c => c.length >= 2).slice(0, 8);

  // 检测结构分段标记
  const isSectionHeader = (text) => /第[一二三\d]+[部分章节]|^[一二三四五六七八九十][、，]|关于[A-Za-z一-鿿]{2,8}/.test(text.trim());

  // 分组
  const groups = [];
  let currentGroup = [];

  for (const s of ranked) {
    const isNewSection = isSectionHeader(s.text);

    if (isNewSection && currentGroup.length > 0) {
      groups.push([...currentGroup]);
      currentGroup = [s];
    } else {
      currentGroup.push(s);
    }
  }
  if (currentGroup.length > 0) groups.push(currentGroup);

  // 后处理：合并只有1句的极小簇
  const finalGroups = [];
  for (let i = 0; i < groups.length; i++) {
    if (groups[i].length <= 1 && finalGroups.length > 0) {
      finalGroups[finalGroups.length - 1].push(...groups[i]);
    } else {
      finalGroups.push([...groups[i]]);
    }
  }

  // 4. 生成输出：每个簇一个主题
  const domainLabels = {
    '沟通与表达': '沟通要点', '时间与效率': '效率方法', '学习与认知': '学习要点',
    '写作与输出': '写作要点', '领导与管理': '管理方法', '产品与设计': '产品思路',
    '技术与编程': '技术要点', '营销与增长': '增长策略', '心理与情商': '心理洞察',
    '数据与分析': '数据洞察', '职场与成长': '职场建议',
  };

  let result = `## 📋 文章摘要\n\n`;
  result += `**原文主题**：${title}\n\n`;

  let themeNum = 1;
  for (const group of groups) {
    if (group.length === 0) continue;
    // 用中心性最高的句子作为主题标题
    const top = group.reduce((a, b) => a.centrality > b.centrality ? a : b);
    const themeTitle = top.text.slice(0, 30).replace(/[，。！？\n]/g, '') + (top.text.length > 30 ? '…' : '');

    result += `### 🔹 ${themeTitle}\n\n`;

    // 簇内去重：相似度 > 0.8 的句子只保留中心性最高的
    const deduped = [];
    const groupVectors = computeTFIDF(group.map(s => s.text), tokenize).sentenceVectors;
    for (let i = 0; i < group.length; i++) {
      let isDuplicate = false;
      for (const kept of deduped) {
        const sim = cosineSimilarity(groupVectors[i], groupVectors[group.indexOf(kept)]);
        if (sim > 0.8) {
          // 保留中心性更高的
          if (group[i].centrality > kept.centrality) {
            deduped.splice(deduped.indexOf(kept), 1);
            isDuplicate = false;
          } else {
            isDuplicate = true;
          }
          break;
        }
      }
      if (!isDuplicate) deduped.push(group[i]);
    }

    for (const s of deduped) {
      const cleaned = s.text.trim().replace(/^[\d\.\s、①②③④⑤\*\-\#]+/, '');
      if (cleaned.length > 5) {
        result += `- ${cleaned}\n`;
      }
    }
    result += '\n';
    themeNum++;
  }

  // 5. 添加关键概念
  if (coreConcepts && coreConcepts.length > 0) {
    result += `**🔑 核心关键词**：${coreConcepts.filter(c => c.length >= 2).slice(0, 6).join(' · ')}\n`;
  }

  return result;
}

// ==================== 模式 2：提取特定信息 ====================

const EXTRACTION_RULES = [
  { category: '💰 价格/金额', patterns: [
    /¥\s*\d[\d,.]*(?:\.\d{1,2})?/g,
    /(?:人民币|售价|价格|费用|成本|单价|总价)[：:]\s*\d[\d,.]*/g,
    /\$\s*\d[\d,.]*/g,
    /\d[\d,.]*\s*(?:元|美元|美金|欧元|日元|英镑|港币)/g,
    /\d[\d,.]*\s*(?:万|亿|千|百)?\s*(?:元)/g,
  ]},
  { category: '📅 日期/时间', patterns: [
    /\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日/g,
    /\d{4}\s*年\s*\d{1,2}\s*月/g,
    /\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/g,
    /\d{1,2}\s*月\s*\d{1,2}\s*日/g,
    /(?:今天|明天|昨天|本周|下周|本月|下月|今年|明年)/g,
  ]},
  { category: '👤 人名/角色', patterns: [
    /(?:作者|发言人|负责人|联系人|CEO|创始人|经理|总监|专家|教授|博士)[：:]\s*[一-鿿A-Za-z·]{2,20}/g,
    /(?:采访|访谈|专访)\s*[一-鿿A-Za-z·]{2,10}/g,
    // 常见姓氏+名结构（约100个常见姓做前缀匹配，减少误报）
    /[王李张刘陈杨黄赵周吴徐孙马胡朱郭何罗高林郑梁谢唐许冯宋韩邓彭曹曾田萧潘袁蔡蒋余于杜叶程魏苏吕丁任卢姚沈钟姜崔谭廖范汪陆金石戴贾韦夏付方白邹孟熊秦邱江尹薛闫段雷侯龙史陶黎贺顾毛郝龚邵万钱严覃武戴莫孔]/g,
  ]},
  { category: '📊 数据/统计', patterns: [
    /\d{1,3}(?:\.\d{1,2})?\s*%/g,
    /(?:增长|下降|提升|降低|增加|减少|上涨|下跌)\s*\d{1,3}(?:\.\d{1,2})?\s*%/g,
    /\d[\d,.]*\s*(?:次|个|件|人|家|项|笔|万|亿)/g,
  ]},
  { category: '📧 联系方式', patterns: [
    /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g,
    /1[3-9]\d{9}/g,
    /https?:\/\/[^\s，。！？\n]+/g,
    /(?:电话|手机|微信|QQ|邮箱)[：:]\s*\S+/g,
  ]},
];

function generateExtraction(analysis, userProfile) {
  const { title, sentences } = analysis;
  if (!sentences || sentences.length === 0) {
    return '内容过短，无法提取信息。';
  }

  const content = sentences.join('\n');
  const results = [];

  for (const rule of EXTRACTION_RULES) {
    for (const pattern of rule.patterns) {
      // 重置 lastIndex（全局正则需要）
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const value = match[0].trim();
        // 过滤太短或纯数字的匹配
        if (value.length < 2) continue;
        // 提取包含该值的上下文片段（前后各取30字）
        let context = '';
        for (let i = 0; i < sentences.length; i++) {
          const idx = sentences[i].indexOf(value);
          if (idx >= 0) {
            const start = Math.max(0, idx - 25);
            const end = Math.min(sentences[i].length, idx + value.length + 25);
            context = sentences[i].slice(start, end);
            if (start > 0) context = '…' + context;
            if (end < sentences[i].length) context = context + '…';
            break;
          }
        }
        if (!context) context = value;
        // 去重
        if (!results.some(r => r.value === value && r.category === rule.category)) {
          results.push({ category: rule.category, value, context });
        }
      }
    }
  }

  // 过滤人名规则产生的大量误匹配
  const filtered = results.filter(r => {
    if (r.category === '👤 人名/角色') {
      const nonNameWords = /[的了着过是就也会要和或在不但与从被把对向让用能可以需要应该这个那个什么怎么一二三四五六七八九十百千万亿年月日时分秒]/;
      return r.value.length >= 2 && r.value.length <= 4 && !nonNameWords.test(r.value) && !/^\d+$/.test(r.value);
    }
    return true;
  });

  // 去子串：同一类别中，"2025年9月20日"保留，"2025年9月"和"9月20日"去掉
  const deduped = filtered.filter((r, i, arr) => {
    // 检查是否有更完整的同类匹配包含当前值
    const hasBetter = arr.some((other, j) =>
      j !== i &&
      other.category === r.category &&
      other.value !== r.value &&
      other.value.includes(r.value) &&
      other.value.length > r.value.length
    );
    return !hasBetter;
  });

  // 限制每类最多5条
  const categoryCount = new Map();
  const limited = deduped.filter(r => {
    const count = categoryCount.get(r.category) || 0;
    if (count >= 5) return false;
    categoryCount.set(r.category, count + 1);
    return true;
  });

  // 无结果时给出提示
  if (limited.length === 0) {
    let result = `## 🔍 信息提取结果\n\n`;
    result += `**原文**：${title}\n\n`;
    result += `未检测到明显的结构化信息（价格、日期、人名、数据等）。\n\n`;
    result += `💡 试试其他模式：切换到"总结重点"获取文章概要，或"能力转化"将其变为行动指南。`;
    return result;
  }

  // 按类别分组输出
  const grouped = new Map();
  for (const r of limited) {
    if (!grouped.has(r.category)) grouped.set(r.category, []);
    grouped.get(r.category).push(r);
  }

  const categoryDescriptions = {
    '💰 价格/金额': '文中涉及的价格、费用、金额信息',
    '📅 日期/时间': '文中提到的时间节点和日期',
    '👤 人名/角色': '文中出现的人物、角色、作者',
    '📊 数据/统计': '文中引用的数字、百分比、统计数据',
    '📧 联系方式': '文中包含的联系电话、邮箱、网址',
  };

  let result = `## 🔍 信息提取结果\n\n`;
  result += `**原文**：${title}\n`;
  result += `**提取概况**：共检测到 **${limited.length}** 条结构化信息，涵盖 ${grouped.size} 个类别。\n\n`;

  for (const [cat, items] of grouped) {
    result += `### ${cat}\n`;
    result += `> ${categoryDescriptions[cat] || '其他信息'}\n\n`;
    result += `| 提取内容 | 上下文 |
|----------|--------|
`;
    for (const item of items) {
      const escapedContext = item.context.replace(/\|/g, '｜').replace(/\n/g, ' ');
      result += `| ${item.value} | ${escapedContext} |
`;
    }
    result += '\n';
  }

  result += `> 💡 以上信息从原文自动提取，建议核对准确性。如需更全面的分析，可切换到"总结重点"或"批判分析"模式。`;

  result += `\n> 共提取 ${limited.length} 条信息。如有遗漏，请切换到其他模式获取更全面的分析。`;
  return result;
}

// ==================== 模式 3：批判分析 ====================

const LOGICAL_FALLACIES = [
  { name: '因果倒置', pattern: /因为.*所以|因此|导致|造成|引起/,
    check: (sentence, context) => {
      // 简化的因果检测：如果句子声称A导致B，检查是否有数据支撑
      const hasData = /\d+%|数据|研究|实验|调查|统计|证明/.test(context);
      return hasData ? null : '该因果断言缺少数据或研究支撑，可能是相关性被误读为因果性';
    }
  },
  { name: '以偏概全', pattern: /所有|全部|每个人|从来|永远|绝对|一定|必然|总是|从不/,
    check: (sentence) => {
      return '使用了绝对化表述，可能存在以偏概全的风险——反例是否存在？';
    }
  },
  { name: '诉诸权威', pattern: /专家|权威|大师|名人|某某说|根据.*研究/,
    check: (sentence) => {
      const hasSpecificSource = /《|》|"\w+|"\w+|某期刊|某大学|某机构/.test(sentence);
      return hasSpecificSource ? null : '引用了"权威"但未指明具体来源，可能是诉诸权威的论证';
    }
  },
  { name: '虚假两难', pattern: /要么.*要么|不是.*就是|非.*即|二选一|只有.*才能/,
    check: (sentence) => {
      return '将复杂问题简化为二元选择，现实中可能存在第三种或更多的可能性';
    }
  },
];

function generateCritique(analysis, userProfile) {
  const { title, sentences, centrality, coreConcepts } = analysis;
  if (!sentences || sentences.length === 0) {
    return '内容过短，无法进行批判分析。';
  }

  const role = userProfile ? userProfile.split(/[,，、]/)[0] : '读者';

  let result = `## 💬 批判分析\n\n`;
  result += `**分析对象**：${title}\n`;
  result += `**分析视角**：以${role}的身份审视本文的论证质量\n\n---\n\n`;

  // 1. 找出所有中心性 > 0.15 且有实质内容的句子
  const allClaims = sentences
    .map((text, idx) => ({ text, idx, centrality: centrality ? centrality[idx] : 0 }))
    .filter(c => c.centrality > 0.15 && c.text.trim().length >= 15)
    .sort((a, b) => b.centrality - a.centrality);

  if (allClaims.length === 0) {
    result += `未找到可分析的明确论点。文章可能为纯描述性或叙述性内容。\n`;
    return result;
  }

  // 2. 去重：相似度过高的只保留中心性最高的
  const deduped = [];
  if (allClaims.length > 1) {
    const claimVecs = computeTFIDF(allClaims.map(c => c.text), tokenize).sentenceVectors;
    for (let i = 0; i < allClaims.length; i++) {
      let dup = false;
      for (const k of deduped) {
        if (cosineSimilarity(claimVecs[i], claimVecs[allClaims.indexOf(k)]) > 0.6) {
          dup = true; break;
        }
      }
      if (!dup) deduped.push(allClaims[i]);
    }
  } else {
    deduped.push(...allClaims);
  }

  // 3. 逐条分析，有多少输出多少
  const counterPerspectives = [
    '如果从相反的角度看，是否存在同样合理的解释？',
    '这个结论是否适用于所有场景，还是仅在特定条件下成立？',
    '如果数据或前提发生变化，这个论点是否仍然成立？',
    '这个观点有没有隐含的前提假设？这些假设本身是否经得起检验？',
    '有没有反例可以挑战这个结论？',
  ];

  for (let i = 0; i < deduped.length; i++) {
    const claim = deduped[i];
    const s = claim.text.trim();

    result += `### 论点 ${i + 1}：${s.slice(0, 60)}${s.length > 60 ? '…' : ''}\n\n`;

    const ctxStart = Math.max(0, claim.idx - 2);
    const ctxEnd = Math.min(sentences.length, claim.idx + 3);
    const context = sentences.slice(ctxStart, ctxEnd).join(' ');

    const hasEvidence = /\d+%|例如|比如|举例|数据|研究|实验|案例|实例|根据|据/.test(context);
    result += `- 📌 **支撑论据**：${hasEvidence ? '原文提供了数据或案例支撑' : '⚠️ 缺乏具体数据或案例支撑'}\n`;

    const fallacies = [];
    for (const f of LOGICAL_FALLACIES) {
      if (f.pattern.test(s)) {
        const issue = f.check(s, context);
        if (issue) fallacies.push(`**${f.name}**：${issue}`);
      }
    }
    if (fallacies.length > 0) {
      result += `- ⚠️ **潜在漏洞**：${fallacies.join('；')}\n`;
    }

    result += `- 🔄 **反方视角**：${counterPerspectives[i % counterPerspectives.length]}\n\n`;
  }

  // 4. 总体评价
  result += `---\n\n`;
  result += `**📊 分析总结**：全文 ${sentences.length} 句，识别出 **${deduped.length}** 个可分析论点。`;
  result += `${role ? '建议' + role : '建议读者'}在采纳本文观点时，重点验证以上论点的实证支撑和边界条件。`;

  return result;
}

// ==================== API 接口 ====================

/**
 * POST /api/parse-file
 * 上传 Word (.docx) 或 Excel (.xlsx) 文件，返回解析后的纯文本
 *
 * 请求: multipart/form-data，字段名 "file"
 * 返回: { text: string, fileName: string, fileType: string, charCount: number }
 */
app.post('/api/parse-file', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '请选择一个文件上传。' });
    }

    const { originalname, mimetype, buffer, size } = req.file;
    const ext = path.extname(originalname).toLowerCase();
    let text = '';

    if (ext === '.docx') {
      text = await parseDocx(buffer);
    } else if (ext === '.xlsx') {
      text = parseXlsx(buffer);
    } else if (ext === '.md' || ext === '.txt') {
      // 先尝试 UTF-8，如果含乱码则用 GBK
      text = buffer.toString('utf-8').trim();
      // 检测是否有乱码特征（大量 � 或不可打印字符）
      const replacementChars = (text.match(/�/g) || []).length;
      if (replacementChars > text.length * 0.05) {
        try {
          const iconv = require('iconv-lite');
          text = iconv.decode(buffer, 'gbk').trim();
        } catch { /* iconv 不可用时保持 UTF-8 结果 */ }
      }
    } else {
      return res.status(400).json({ error: `不支持的文件格式 "${ext}"，请上传 .docx、.xlsx、.md 或 .txt 文件。` });
    }

    if (!text || text.length === 0) {
      return res.status(400).json({ error: '文件解析结果为空，请检查文件内容。' });
    }

    const typeMap = { '.docx': 'Word 文档', '.xlsx': 'Excel 表格', '.md': 'Markdown 文档' };
    res.json({
      text,
      fileName: originalname,
      fileType: typeMap[ext] || '文档',
      charCount: text.length,
    });
  } catch (err) {
    // multer 的文件类型错误
    if (err.message && err.message.includes('仅支持')) {
      return res.status(400).json({ error: err.message });
    }
    console.error('[parse-file] 错误:', err.message);
    res.status(500).json({ error: `文件解析失败: ${err.message}` });
  }
});

/**
 * POST /api/fetch-url
 * 抓取网页/社交媒体链接的内容，提取文案
 *
 * 请求体：{ url: string }
 * 返回：{ text: string, title: string, source: string, charCount: number }
 */
app.post('/api/fetch-url', async (req, res) => {
  try {
    const { url } = req.body;

    if (!url || !url.trim()) {
      return res.status(400).json({ error: '请提供一个链接。' });
    }

    let parsedUrl;
    try {
      parsedUrl = new URL(url.trim());
    } catch {
      return res.status(400).json({ error: '链接格式不正确，请输入完整的网址（以 http:// 或 https:// 开头）。' });
    }

    if (!['http:', 'https:'].includes(parsedUrl.protocol)) {
      return res.status(400).json({ error: '仅支持 http:// 和 https:// 链接。' });
    }

    // 抓取网页
    const html = await fetchUrl(parsedUrl);

    if (!html || html.trim().length === 0) {
      return res.status(400).json({ error: '无法读取该链接的内容，请确认链接可访问。' });
    }

    // 从 HTML 中提取文本
    let { title, text, source } = extractTextFromHtml(html, url);

    // 如果提取的内容太少，尝试平台专用 API
    if ((!text || text.trim().length < 20) && source !== 'example.com') {
      const apiResult = await tryPlatformApi(url, source);
      if (apiResult) {
        title = apiResult.title || title;
        text = apiResult.text;
        source = apiResult.source || source;
      }
    }

    if (!text || text.trim().length < 10) {
      return res.status(400).json({
        error: '未能从该链接提取到有效文本（可能页面需要 JavaScript 渲染，或为纯图片/视频内容）。请尝试直接粘贴文案。',
      });
    }

    res.json({
      text: text.trim(),
      title: title || '未提取到标题',
      source,
      charCount: text.length,
    });
  } catch (err) {
    console.error('[fetch-url] 错误:', err.message);
    res.status(500).json({ error: `链接抓取失败: ${err.message}` });
  }
});

/**
 * HTTP/HTTPS 请求封装
 */
function fetchUrl(parsedUrl) {
  return new Promise((resolve, reject) => {
    const mod = parsedUrl.protocol === 'https:' ? https : http;
    const req = mod.get(
      parsedUrl.href,
      {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
          'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        },
        timeout: 10000,
      },
      (response) => {
        // 处理重定向
        if ([301, 302, 307, 308].includes(response.statusCode)) {
          const redirectUrl = response.headers.location;
          if (redirectUrl) {
            try {
              const redirectParsed = new URL(redirectUrl, parsedUrl.href);
              return fetchUrl(redirectParsed).then(resolve).catch(reject);
            } catch {
              reject(new Error('重定向地址无效'));
              return;
            }
          }
        }

        if (response.statusCode !== 200) {
          reject(new Error(`服务器返回 ${response.statusCode}`));
          return;
        }

        const chunks = [];
        response.on('data', (chunk) => chunks.push(chunk));
        response.on('end', () => {
          const html = Buffer.concat(chunks).toString('utf-8');
          resolve(html);
        });
      }
    );

    req.on('error', (err) => reject(new Error(`网络请求失败: ${err.message}`)));
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('请求超时（10秒）'));
    });
  });
}

/**
 * 从 HTML 中提取文本内容
 *
 * 策略：
 * 1. 优先提取 Open Graph / meta 描述
 * 2. 去掉 script/style/noscript/iframe 等无用标签
 * 3. 提取 body 中的可见文本
 * 4. 清理多余空白
 */
function extractTextFromHtml(html, sourceUrl) {
  let title = '';
  let description = '';
  let source = '';

  // ===== 1. 尽力提取标题 =====
  // Open Graph 标题
  let m = html.match(/<meta[^>]+property="og:title"[^>]+content="([^"]+)"/i);
  if (m) title = m[1];
  // Twitter 标题
  if (!title) {
    m = html.match(/<meta[^>]+(?:name|property)="twitter:title"[^>]+content="([^"]+)"/i);
    if (m) title = m[1];
  }
  // HTML <title> 标签
  if (!title) {
    m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    if (m) title = m[1].trim().replace(/\s+/g, ' ');
  }

  // ===== 2. 尽力提取描述 =====
  m = html.match(/<meta[^>]+name="description"[^>]+content="([^"]+)"/i);
  if (m) description = m[1];
  if (!description) {
    m = html.match(/<meta[^>]+property="og:description"[^>]+content="([^"]+)"/i);
    if (m) description = m[1];
  }
  if (!description) {
    m = html.match(/<meta[^>]+(?:name|property)="twitter:description"[^>]+content="([^"]+)"/i);
    if (m) description = m[1];
  }

  // ===== 3. 从 JSON-LD / __NEXT_DATA__ / __INITIAL_STATE__ 中挖掘内容 =====
  let jsonData = '';

  // 3a. 多个 <script type="application/ld+json">
  const ldJsonMatches = html.match(/<script[^>]+type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/gi);
  if (ldJsonMatches) {
    for (const match of ldJsonMatches) {
      const jsonStr = match.replace(/<[^>]+>/g, '').trim();
      try {
        const ld = JSON.parse(jsonStr);
        const flat = JSON.stringify(ld);
        jsonData += flat + '\n';
      } catch { /* JSON 不完整，跳过 */ }
    }
  }

  // 3b. __NEXT_DATA__ / __INITIAL_STATE__ 等 SSR 数据注入
  const ssrPatterns = [
    /__NEXT_DATA__\s*=\s*({[\s\S]*?});/,
    /window\.__INITIAL_STATE__\s*=\s*({[\s\S]*?});/,
    /__NUXT__\s*=\s*({[\s\S]*?});/,
    /window\.__DATA__\s*=\s*({[\s\S]*?});/,
    /window\.__PRELOADED_STATE__\s*=\s*({[\s\S]*?});/,
    /window\._SSR_DATA_\s*=\s*({[\s\S]*?});/,
  ];
  for (const pattern of ssrPatterns) {
    const m = html.match(pattern);
    if (m) {
      try {
        const parsed = JSON.parse(m[1]);
        jsonData += JSON.stringify(parsed) + '\n';
      } catch { /* 截断的 JSON */ }
    }
  }

  // 3c. 从 JSON 数据中提取可读文本
  if (jsonData) {
    // 提取所有 "title" "desc" "content" "text" "name" "summary" 字段的值
    const textFields = [];
    const fieldPatterns = [
      /"title"\s*:\s*"([^"]{2,200})"/g,
      /"desc(?:ription)?"\s*:\s*"([^"]{2,500})"/g,
      /"content"\s*:\s*"([^"]{2,500})"/g,
      /"text"\s*:\s*"([^"]{2,500})"/g,
      /"summary"\s*:\s*"([^"]{2,500})"/g,
      /"name"\s*:\s*"([^"]{2,100})"/g,
      /"caption"\s*:\s*"([^"]{2,500})"/g,
      /"headline"\s*:\s*"([^"]{2,200})"/g,
      /"bio"\s*:\s*"([^"]{2,500})"/g,
    ];
    for (const pattern of fieldPatterns) {
      let fm;
      while ((fm = pattern.exec(jsonData)) !== null) {
        const val = fm[1].replace(/\\"/g, '"').replace(/\\n/g, '\n').replace(/\\t/g, ' ');
        if (val.length >= 2 && !textFields.includes(val)) {
          textFields.push(val);
        }
      }
    }

    // 如果 JSON 中挖到了标题/描述，优先使用
    if (!title && textFields.length > 0) {
      title = textFields.find(t => t.length >= 4 && t.length <= 100) || title;
    }
    if (!description && textFields.length > 1) {
      const desc = textFields.find(t => t.length > 20 && t !== title);
      if (desc) description = desc;
    }

    // 将 JSON 中挖到的所有文本拼接为补充内容
    if (textFields.length > 0) {
      jsonData = textFields.join('\n');
    }
  }

  // ===== 4. 识别来源平台 =====
  const hostname = new URL(sourceUrl).hostname.replace('www.', '');
  if (hostname.includes('douyin.com') || hostname.includes('iesdouyin.com')) {
    source = '抖音';
  } else if (hostname.includes('xiaohongshu.com') || hostname.includes('xhslink.com')) {
    source = '小红书';
  } else if (hostname.includes('zhihu.com')) {
    source = '知乎';
  } else if (hostname.includes('weixin.qq.com') || hostname.includes('mp.weixin.qq.com')) {
    source = '微信公众号';
  } else if (hostname.includes('bilibili.com') || hostname.includes('b23.tv')) {
    source = 'B站';
  } else if (hostname.includes('weibo.com')) {
    source = '微博';
  } else {
    source = hostname;
  }

  // ===== 5. 提取 body 中的可见文本 =====
  let bodyText = html;
  // 去掉 head
  bodyText = bodyText.replace(/<head[\s\S]*?<\/head>/gi, '');
  // 去掉 script/style/noscript/iframe/svg/nav/footer/header
  bodyText = bodyText.replace(/<(script|style|noscript|iframe|svg|nav|footer|header)[\s\S]*?<\/\1>/gi, '');
  // 去掉 HTML 注释
  bodyText = bodyText.replace(/<!--[\s\S]*?-->/g, '');
  // 去掉所有 HTML 标签
  bodyText = bodyText.replace(/<[^>]+>/g, '\n');
  // 解码 HTML 实体
  bodyText = bodyText.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#x27;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCharCode(d))
    .replace(/&nbsp;/g, ' ');
  // 清理冗余空白
  bodyText = bodyText.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();

  // 过滤 UI 噪音行（导航、按钮、广告等）
  const noisePatterns = [
    /^(登录|注册|退出登录|首页|搜索|关注|分享|评论|点赞|收藏|下载APP?|打开|扫一扫|查看更多|阅读原文|广告|推广|菜单|导航|返回|投稿|反馈|设置|我的|消息|通知|发布|上传)$/,
    /^(关注|投币|收藏|分享|充电|弹幕|稍后再看|缓存|举报|不感兴趣)$/, // B站
    /^(阅读|赞|在看|分享|留言|写留言|精选留言|作者回复)$/, // 公众号
    /^(赞同|反对|喜欢|收藏|评论|分享|举报|没有帮助|作者|发布于|编辑于)$/, // 知乎
    /^(小红书|笔记详情|相关推荐|你可能感兴趣|热门搜索)$/, // 小红书
    /^(CopyRight|©|版权所有|ICP备|公网安备|友情链接|关于我们|联系我们|用户协议|隐私政策|营业执照)$/,
    /^(关注我们|扫码关注|长按识别|商务合作|广告投放|人才招聘|帮助中心|意见反馈)$/,
  ];
  bodyText = bodyText.split('\n')
    .filter(line => {
      const trimmed = line.trim();
      if (trimmed.length <= 1) return false;
      for (const pattern of noisePatterns) {
        if (pattern.test(trimmed)) return false;
      }
      return true;
    })
    .join('\n');

  // ===== 6. 拼接最终结果 =====
  const parts = [];
  if (title) parts.push(`【标题】${title}`);
  if (description) parts.push(`【简介】${description}`);

  // JSON 数据中挖到的文本（优先于 body 中的噪音）
  if (jsonData && typeof jsonData === 'string' && jsonData.trim().length > 20) {
    parts.push(`【内容】\n${jsonData}`);
  }

  // body 文本作为补充
  if (bodyText.length > 10) {
    // 如果已有 JSON 内容，body 只取前 2000 字作为补充
    const bodyPart = jsonData ? bodyText.slice(0, 2000) : bodyText;
    if (parts.length === 0 || bodyPart.trim().length > 20) {
      parts.push(bodyPart);
    }
  }

  const finalText = parts.join('\n\n').trim();

  // ===== 7. 如果最终结果仍然很短，告知用户 =====
  if (finalText.length < 15) {
    const tips = {
      '抖音': '抖音视频页面需要 JavaScript 渲染，无法直接抓取。\n💡 建议：在抖音 App 内复制视频文案，或使用抖音分享链接中的"复制链接"后粘贴到此处。\n💡 也可以直接在下方文本框粘贴视频的标题和描述文字。',
      '小红书': '小红书笔记页面可能需要登录才能查看完整内容。\n💡 建议：在小红书 App 内复制笔记文字，直接粘贴到下方文本框。',
      'B站': 'B站视频页面可能需要 JavaScript 渲染。\n💡 建议：复制视频简介文字直接粘贴，或使用 BV 号在第三方工具中提取。',
    };
    const tip = tips[source] || `${source} 页面可能使用了 JavaScript 动态加载内容，服务器无法直接抓取。\n💡 建议：直接复制原文粘贴到下方文本框中。`;

    return {
      title: title || '未提取到标题',
      text: `【⚠️ 自动抓取失败】\n${tip}`,
      source,
    };
  }

  return {
    title: title || '未提取到标题',
    text: finalText.slice(0, 10000), // 最多 10000 字
    source,
  };
}

// ==================== 平台专用 API 降级策略 ====================

/**
 * 当通用 HTML 抓取拿不到内容时，尝试调用平台公开 API
 */
async function tryPlatformApi(url, detectedSource) {
  const hostname = new URL(url).hostname.replace('www.', '');

  // B站：api.bilibili.com
  if (hostname.includes('bilibili.com') || hostname.includes('b23.tv')) {
    // 从 URL 中提取 BV 号或 aid
    const bvMatch = url.match(/BV[a-zA-Z0-9]{10}/);
    const avMatch = url.match(/av(\d+)/i);
    let apiUrl;
    if (bvMatch) {
      apiUrl = `https://api.bilibili.com/x/web-interface/view?bvid=${bvMatch[0]}`;
    } else if (avMatch) {
      apiUrl = `https://api.bilibili.com/x/web-interface/view?aid=${avMatch[1]}`;
    } else {
      return null;
    }
    return await fetchBilibiliApi(apiUrl);
  }

  // 抖音：iesdouyin.com API（公开接口）
  if (hostname.includes('douyin.com') || hostname.includes('iesdouyin.com')) {
    // 从各种 URL 格式中提取视频 ID
    const idPatterns = [
      /video\/(\d+)/,
      /note\/(\d+)/,
      /modal_id=(\d+)/,
      /aweme_id=(\d+)/,
      /item_id=(\d+)/,
    ];
    let videoId = null;
    for (const p of idPatterns) {
      const m = url.match(p);
      if (m) { videoId = m[1]; break; }
    }
    if (!videoId) return null;
    return await fetchDouyinApi(videoId);
  }

  // 知乎：直接抓取通常够用，不需要降级
  return null;
}

async function fetchBilibiliApi(apiUrl) {
  try {
    const parsed = new URL(apiUrl);
    const html = await new Promise((resolve, reject) => {
      https.get(apiUrl, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
          'Referer': 'https://www.bilibili.com/',
        },
        timeout: 8000,
      }, (res) => {
        let data = '';
        res.on('data', c => data += c);
        res.on('end', () => resolve(data));
      }).on('error', reject).on('timeout', function() { this.destroy(); resolve(''); });
    });

    if (!html) return null;
    const json = JSON.parse(html);
    if (json.code !== 0 || !json.data) return null;

    const d = json.data;
    const parts = [];
    if (d.title) parts.push(`【标题】${d.title}`);
    if (d.desc) parts.push(`【简介】${d.desc}`);
    if (d.dynamic) parts.push(`【动态】${d.dynamic}`);

    if (parts.length > 0) {
      return { title: d.title || '', text: parts.join('\n\n'), source: 'B站' };
    }
  } catch { /* API 不可用 */ }
  return null;
}

async function fetchDouyinApi(videoId) {
  try {
    const apiUrl = `https://www.iesdouyin.com/web/api/v2/aweme/iteminfo/?item_ids=${videoId}`;
    const html = await new Promise((resolve, reject) => {
      https.get(apiUrl, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
          'Referer': 'https://www.douyin.com/',
        },
        timeout: 8000,
      }, (res) => {
        let data = '';
        res.on('data', c => data += c);
        res.on('end', () => resolve(data));
      }).on('error', reject).on('timeout', function() { this.destroy(); resolve(''); });
    });

    if (!html) return null;
    const json = JSON.parse(html);
    if (json.status_code !== 0 || !json.item_list || json.item_list.length === 0) return null;

    const item = json.item_list[0];
    const parts = [];
    if (item.desc) parts.push(`【标题】${item.desc}`);
    if (item.author?.nickname) parts.push(`【作者】${item.author.nickname}`);
    // 话题标签
    if (item.text_extra) {
      const tags = item.text_extra.map(t => `#${t.hashtag_name}`).join(' ');
      if (tags) parts.push(`【话题】${tags}`);
    }
    // 音乐
    if (item.music?.title) parts.push(`【音乐】${item.music.title}`);

    if (parts.length > 0) {
      return { title: item.desc || '', text: parts.join('\n\n'), source: '抖音' };
    }
  } catch { /* API 不可用 */ }
  return null;
}

/**
 * POST /api/transform
 * 接收用户内容和转化模式，输出对应格式的结果
 *
 * mode 参数：
 * - 'summarize'   → 总结文章重点（结构化摘要）
 * - 'extract'     → 提取特定信息（价格/日期/人名/数据）
 * - 'capability'  → 能力转化（能力卡片+行动脚本+角色扮演）
 * - 'critique'    → 批判分析（论点审视+逻辑漏洞+反方视角）
 * 默认：'capability'（向后兼容）
 */
app.post('/api/transform', async (req, res) => {
  try {
    const { content, userProfile, identity, mode, modes } = req.body;

    if (!content || content.trim().length === 0) {
      return res.status(400).json({ error: '请提供需要转化的内容。' });
    }

    // 限制内容长度
    const MAX_CONTENT_LENGTH = 50000;
    if (content.length > MAX_CONTENT_LENGTH) {
      return res.status(400).json({
        error: `内容过长（${content.length.toLocaleString()} 字符），最多支持 ${MAX_CONTENT_LENGTH.toLocaleString()} 字符。`,
      });
    }

    // 共用分析层：所有模式都先做 TextRank + TF-IDF 分析
    const profile = identity || userProfile || '';
    const analysis = analyzeContent(content, profile);

    // 多模式请求
    const requestModes = (modes && modes.length > 0) ? modes : [mode || 'capability'];

    // 如果只有一个模式，走原有单模式流程
    if (requestModes.length === 1) {
      const selectedMode = requestModes[0];
      let reply = '';
      let conversationHistory = [];

      switch (selectedMode) {
        case 'summarize':
          reply = generateSummary(analysis, profile);
          conversationHistory = [
            { role: 'user', content: `总结以下内容：${content.slice(0, 100)}…` },
            { role: 'assistant', content: reply },
          ];
          break;
        case 'extract':
          reply = generateExtraction(analysis, profile);
          conversationHistory = [
            { role: 'user', content: `提取以下内容中的关键信息：${content.slice(0, 100)}…` },
            { role: 'assistant', content: reply },
          ];
          break;
        case 'critique':
          reply = generateCritique(analysis, profile);
          conversationHistory = [
            { role: 'user', content: `批判分析以下内容：${content.slice(0, 100)}…` },
            { role: 'assistant', content: reply },
          ];
          break;
        case 'capability':
        default:
          const cardData = {
            oneLiner: inferOneLiner(analysis),
            abilityName: inferAbilityName(analysis),
            logic: inferLogic(analysis),
            pitfalls: inferPitfalls(analysis),
          };
          reply = [
            generateCapabilityCard(analysis, profile, cardData),
            generateActionScript(analysis, profile, cardData),
            generateRoleplayInvitation(analysis, profile),
          ].join('\n');
          conversationHistory = [
            { role: 'user', content: profile || '学习者' },
            { role: 'assistant', content: reply },
          ];
          break;
      }

      const sessionId = createSession(analysis, profile);
      res.json({ reply, conversationHistory, sessionId, mode: selectedMode });
    } else {
      // 多模式：一次分析，多管道生成
      const results = {};
      for (const m of requestModes) {
        switch (m) {
          case 'summarize': results.summarize = generateSummary(analysis, profile); break;
          case 'extract': results.extract = generateExtraction(analysis, profile); break;
          case 'capability': {
            const cd = { oneLiner: inferOneLiner(analysis), abilityName: inferAbilityName(analysis), logic: inferLogic(analysis), pitfalls: inferPitfalls(analysis) };
            results.capability = [generateCapabilityCard(analysis, profile, cd), generateActionScript(analysis, profile, cd), generateRoleplayInvitation(analysis, profile)].join('\n');
            break;
          }
          case 'critique': results.critique = generateCritique(analysis, profile); break;
        }
      }

      const sessionId = createSession(analysis, profile);
      const conversationHistory = [
        { role: 'user', content: profile || '学习者' },
        { role: 'assistant', content: JSON.stringify(results) },
      ];
      res.json({ results, conversationHistory, sessionId, modes: requestModes });
    }
  } catch (err) {
    console.error('[transform] 错误:', err.message);
    res.status(500).json({ error: err.message });
  }
});

/**
 * POST /api/roleplay
 * 继续角色扮演对话，给出点评 + 收集反馈
 */
app.post('/api/roleplay', async (req, res) => {
  try {
    const { conversationHistory, userReply, sessionId } = req.body;

    if (!userReply || userReply.trim().length === 0) {
      return res.status(400).json({ error: '请输入你的角色扮演回复。' });
    }

    if (!conversationHistory || conversationHistory.length === 0) {
      return res.status(400).json({ error: '请先完成内容转化。' });
    }

    // 从服务端会话中找回分析结果
    const session = sessionId ? getSession(sessionId) : null;
    const analysis = session ? session.analysis : null;

    // 生成点评
    const reply = generateRoleplayFeedback(userReply, analysis);

    // 更新会话历史
    const updatedHistory = [
      ...conversationHistory,
      { role: 'user', content: userReply },
      { role: 'assistant', content: reply },
    ];

    res.json({ reply, conversationHistory: updatedHistory });
  } catch (err) {
    console.error('[roleplay] 错误:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ==================== QR Code 生成 ====================

const QRCode = require('qrcode');

/**
 * GET /api/qrcode?url=xxx
 * 服务端生成二维码 PNG，不依赖任何外部 CDN/API
 */
app.get('/api/qrcode', async (req, res) => {
  try {
    const url = req.query.url;
    if (!url) return res.status(400).json({ error: '请提供 url 参数' });

    const pngBuffer = await QRCode.toBuffer(url, {
      width: 250,
      margin: 2,
      color: { dark: '#1e1b2e', light: '#ffffff' },
    });

    res.setHeader('Content-Type', 'image/png');
    res.setHeader('Cache-Control', 'public, max-age=3600');
    res.send(pngBuffer);
  } catch (err) {
    res.status(500).json({ error: '二维码生成失败' });
  }
});

// ==================== 导出与启动 ====================

// 导出 Express 应用（供 Vercel serverless 使用）
module.exports = app;

// 本地开发时启动 HTTP 服务
// require.main === module 表示直接运行 node server.js（而非被其他模块引用）
if (require.main === module) {
  const server = app.listen(PORT, () => {
    console.log(`🧠 Capable 服务已启动（离线模式）→ http://localhost:${PORT}`);
  });

  // 优雅关闭：在收到终止信号时关闭 HTTP 服务器并清理会话
  function gracefulShutdown(signal) {
    console.log(`\n收到 ${signal} 信号，正在关闭服务…`);
    server.close(() => {
      console.log('HTTP 服务已关闭');
      sessions.clear();
      process.exit(0);
    });
    // 5 秒后强制退出
    setTimeout(() => {
      console.error('强制退出（超时）');
      process.exit(1);
    }, 5000);
  }

  process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));
  process.on('SIGINT', () => gracefulShutdown('SIGINT'));
}
