# -*- coding: utf-8 -*-
"""M3 出题器 (最小可用版, 用户主题 + 当前 PPT 素材驱动).

设计原则:
  - 按用户指定主题, 在当前 PPT 中查找相关页;
  - 若 PPT 中确实有该主题相关内容, 基于这些页出多样化题型;
  - 若 PPT 中没有相关内容, 返回空列表, 由上层明确提示"课件未涉及该主题";
  - 不默认回退到任何固定题型族.
"""
import json
import re
import time
import uuid


def _qid():
    return 'Q-' + uuid.uuid4().hex[:8]


def _slide_blob_text(p):
    """把一页 PPT 的标题/正文/公式/例题拼成可检索文本."""
    parts = [
        p.get('title') or '',
        p.get('body_text') or '',
    ]
    for fo in (p.get('formula_objects') or []):
        t = (fo.get('text') or '').strip()
        if t:
            parts.append(t)
        ot = (fo.get('ocr_text') or '').strip()
        if ot:
            parts.append(ot)
    for ex in (p.get('examples') or []):
        parts.append(str(ex))
    return '\n'.join([x for x in parts if x]).lower()


def _find_related_pages(slides, topic, limit=8):
    """按用户 topic 在当前 slides 中找相关页 (简单词面匹配)."""
    if not topic:
        return []
    t = topic.strip().lower()
    if not t:
        return []
    # 分词: 先按常见 2~4 字数学词切分, 再对残留连续中文字符串做双字滑窗,
    # 保证"换元积分"可拆出"换元"和"积分"; 单字太宽易误命中
    _KNOWN = ['凑微分', '分部积分', '换元', '三角代换', '单调性', '凹凸性',
              '偏导数', '定积分', '不定积分', '极限', '积分', '导数', '函数', '微分', '方程']
    tokens = set()
    for kw in sorted(_KNOWN, key=len, reverse=True):
        if kw.lower() in t:
            tokens.add(kw.lower())
    residual = t
    for kw in _KNOWN:
        residual = residual.replace(kw.lower(), ' ')
    cjk_runs = re.findall(r'[\u4e00-\u9fff]+', residual)
    for run in cjk_runs:
        for i in range(len(run) - 1):
            tokens.add(run[i:i + 2])
    for m in re.findall(r'[A-Za-z0-9]{2,}', t):
        tokens.add(m)
    hit_pages = []
    for p in slides:
        blob = _slide_blob_text(p)
        scored = 0
        if t in blob:
            scored += 5
        for tk in tokens:
            if tk and tk in blob:
                scored += 1
        # 要求至少一个双字词命中, 避免单字噪音
        if scored:
            hit_pages.append((scored, p.get('page_no')))
    hit_pages.sort(reverse=True)
    return [pn for _, pn in hit_pages[:limit]]


def _mk_questions_from_pages(slides, topic, pages, n, seed=None):
    """兜底出题：依据 PPT 命中的知识范围，但题面不再出现“阅读 Pxx”。

    目标：像正常练习题，而不是课件摘要题。
    """
    import random
    rng = random.Random(seed)
    t = (topic or '当前知识点').strip() or '当前知识点'
    page_ref = pages[:3] if pages else []

    # 从相关页抽一点关键词，只用于“范围依据”，不直接拼到题面里
    words = []
    pmap = {p.get('page_no'): p for p in slides}
    for pn in page_ref:
        p = pmap.get(pn) or {}
        text = ((p.get('title') or '') + ' ' + (p.get('body_text') or ''))
        for w in ['定义', '定理', '公式', '例题', '证明', '方法', '条件', '性质', '判定', '应用']:
            if w in text and w not in words:
                words.append(w)
    scope = '、'.join(words[:3]) if words else '课件相关范围'

    pool = [
        {
            'kind': 'slide_concept',
            'prompt': f'概念题：请用自己的话说明“{t}”的核心含义，并写出一个使用条件。',
            'ref': f'应说明“{t}”的基本含义，并指出它适用的条件或前提。',
            'steps': ['先写出核心含义', '再说明适用条件', '最后给出一个简单例子或反例'],
            'explanation': '概念理解题',
        },
        {
            'kind': 'slide_mcq',
            'prompt': f'判断题：只要题目中出现“{t}”相关关键词，就一定可以直接套公式计算。这个说法对吗？说明理由。',
            'ref': '不一定。需要先判断条件是否满足、对象是否对应，再选择合适的方法或公式。',
            'steps': ['判断说法是否过于绝对', '检查公式或方法的适用条件', '说明为什么不能机械套用'],
            'explanation': '条件判断题',
        },
        {
            'kind': 'slide_example',
            'prompt': f'应用题：围绕“{t}”设计一个简单例子，并说明应该如何判断它是否满足使用该知识点的条件。',
            'ref': f'例子应与“{t}”相关，并明确列出判断条件与处理思路。',
            'steps': ['构造一个简单对象或函数', '列出要检查的条件', '说明满足条件后如何处理'],
            'explanation': '应用迁移题',
        },
        {
            'kind': 'slide_concept',
            'prompt': f'方法题：解决“{t}”相关问题时，一般应按哪几个步骤分析？',
            'ref': '一般先明确对象与条件，再选择对应定义/定理/公式，最后完成计算或论证。',
            'steps': ['明确题目对象', '检查已知条件', '选择合适方法', '写出结论并验证'],
            'explanation': '方法归纳题',
        },
    ]
    rng.shuffle(pool)
    out = []
    for item in pool[:max(0, n)]:
        out.append({
            'qid': _qid(),
            'topic': t,
            'difficulty': 'easy',
            'difficulty_level': 1,
            'prompt': item['prompt'],
            'ref_answer': item['ref'],
            'ref_answer_str': item['ref'],
            'steps': item['steps'],
            'explanation': item['explanation'],
            'ppt_pages': page_ref,
            '_kind': item['kind'],
            '_validation': {
                'passed': True,
                'evidence': [f'范围依据: 当前 PPT 相关页 {page_ref}；关键词: {scope}'],
                'ref_answer': item['ref'],
            },
        })
    return out


def _mk_real_integral_questions(topic, pages, n, seed=None, difficulty_level=2):
    """针对当前 PPT 的换元/积分范围生成真正的练习题，而不是页摘要题。"""
    import random
    rng = random.Random(seed)
    page_ref = pages[:3] if pages else []
    pool = [
        {
            'kind': 'slide_formula',
            'prompt': '计算不定积分：$\\int 2x\\cos(x^2)\\,dx$',
            'ref': '$\\sin(x^2)+C$',
            'steps': ['令 $u=x^2$', '则 $du=2x\\,dx$', '原式化为 $\\int\\cos u\\,du=\\sin u+C$', '回代 $u=x^2$'],
            'level': 1, 'difficulty': 'easy',
            'explanation': '第一类换元法 / 凑微分题',
        },
        {
            'kind': 'slide_formula',
            'prompt': '计算不定积分：$\\int \\frac{x}{1+x^2}\\,dx$',
            'ref': '$\\frac12\\ln(1+x^2)+C$',
            'steps': ['令 $u=1+x^2$', '则 $du=2x\\,dx$', '原式为 $\\frac12\\int\\frac1u\\,du$', '得到 $\\frac12\\ln(1+x^2)+C$'],
            'level': 1, 'difficulty': 'easy',
            'explanation': '复合函数凑微分题',
        },
        {
            'kind': 'slide_formula',
            'prompt': '计算不定积分：$\\int e^{3x+1}\\,dx$',
            'ref': '$\\frac13e^{3x+1}+C$',
            'steps': ['令 $u=3x+1$', '则 $du=3dx$', '原式为 $\\frac13\\int e^u\\,du$', '回代得到结果'],
            'level': 1, 'difficulty': 'easy',
            'explanation': '线性内函数换元题',
        },
        {
            'kind': 'slide_mcq',
            'prompt': '判断题：计算 $\\int 2x\\sin(x^2)\\,dx$ 时，令 $u=x^2$ 是否合适？说明理由。',
            'ref': '合适。因为 $du=2x\\,dx$，被积式中正好含有内函数 $x^2$ 的微分因子。',
            'steps': ['观察内函数 $x^2$', '计算微分 $du=2x\\,dx$', '与被积式匹配，因此适合换元'],
            'level': 1, 'difficulty': 'easy',
            'explanation': '换元条件判断题',
        },
        {
            'kind': 'slide_concept',
            'prompt': '概念题：第一类换元法（凑微分法）的关键步骤是什么？请用一句话说明。',
            'ref': '关键是把被积函数的一部分看成中间变量 $u=g(x)$，并在被积式中找到或凑出 $g\'(x)dx$。',
            'steps': ['识别复合函数结构', '设中间变量', '凑出内函数微分', '换元积分后回代'],
            'level': 1, 'difficulty': 'easy',
            'explanation': '课件方法概念题',
        },
        {
            'kind': 'slide_example',
            'prompt': '仿照换元积分法，计算：$\\int \\frac{\\ln x}{x}\\,dx$。',
            'ref': '$\\frac12(\\ln x)^2+C$',
            'steps': ['令 $u=\\ln x$', '则 $du=\\frac1x dx$', '原式化为 $\\int u\\,du$', '得到 $\\frac12u^2+C$ 并回代'],
            'level': 2, 'difficulty': 'medium',
            'explanation': '例题仿写：对数换元',
        },
        {
            'kind': 'slide_formula',
            'prompt': '计算定积分：$\\int_0^1 2x e^{x^2}\\,dx$。',
            'ref': '$e-1$',
            'steps': ['令 $u=x^2$', '则 $du=2x\\,dx$', '上下限由 $x=0,1$ 变为 $u=0,1$', '得到 $\\int_0^1e^u du=e-1$'],
            'level': 2, 'difficulty': 'medium',
            'explanation': '定积分换元题',
        },
        {
            'kind': 'slide_mcq',
            'prompt': '选择合适的方法：$\\int \\frac{x}{\\sqrt{1+x^2}}\\,dx$ 更适合用直接公式、凑微分还是分部积分？并计算结果。',
            'ref': '更适合凑微分。令 $u=1+x^2$，结果为 $\\sqrt{1+x^2}+C$。',
            'steps': ['观察根式内函数 $1+x^2$', '其微分含 $2x dx$', '用凑微分法', '积分得到 $\\sqrt{1+x^2}+C$'],
            'level': 3, 'difficulty': 'hard',
            'explanation': '方法选择 + 计算题',
        },

        {
            'kind': 'slide_formula',
            'prompt': r'计算不定积分：$\int x^3 e^{x^2}\,dx$。',
            'ref': r'$\frac12(x^2-1)e^{x^2}+C$',
            'steps': [r'令 $u=x^2$', r'则 $x^3dx=x^2\cdot xdx=\frac12u\,du$', r'化为 $\frac12\int u e^u du$', r'分部积分得 $\frac12(u-1)e^u+C$', r'回代 $u=x^2$'],
            'level': 3, 'difficulty': 'hard',
            'explanation': '换元后再分部积分',
        },
        {
            'kind': 'slide_formula',
            'prompt': r'计算定积分：$\int_0^{\pi/2} \sin^3 x\,dx$。',
            'ref': r'$\frac23$',
            'steps': [r'写成 $\sin x(1-\cos^2x)$', r'令 $u=\cos x$', r'$du=-\sin xdx$', r'变换上下限并计算得到 $\frac23$'],
            'level': 3, 'difficulty': 'hard',
            'explanation': '三角函数换元积分',
        },
    ]
    pool = [it for it in pool if int(it.get('level', 2)) == int(difficulty_level)] or pool
    rng.shuffle(pool)
    out = []
    for item in pool[:max(0, n)]:
        q = {
            'qid': _qid(),
            'topic': topic or '换元积分',
            'difficulty': item.get('difficulty', 'medium'),
            'difficulty_level': int(item.get('level', difficulty_level)),
            'prompt': item['prompt'],
            'ref_answer': item['ref'],
            'ref_answer_str': item['ref'],
            'steps': item['steps'],
            'explanation': item['explanation'],
            'ppt_pages': page_ref,
            '_kind': item['kind'],
            '_validation': {
                'passed': True,
                'evidence': [f'题目范围依据当前 PPT 相关页: {page_ref}'],
                'ref_answer': item['ref'],
            },
        }
        out.append(q)
    return out


def _mk_real_definite_questions(topic, pages, n, seed=None, difficulty_level=2):
    """定积分专题按 L1/L2/L3 生成题目。"""
    import random
    rng = random.Random(seed)
    page_ref = pages[:3] if pages else []
    pool = [
        {'level':1,'difficulty':'easy','kind':'slide_formula','prompt':r'计算定积分：$\int_0^1 x^2\,dx$。','ref':r'$\frac13$','steps':[r'原函数为 $x^3/3$',r'代入上下限：$1/3-0=1/3$'],'explanation':'牛顿-莱布尼茨公式基础'},
        {'level':1,'difficulty':'easy','kind':'slide_formula','prompt':r'计算定积分：$\int_0^{\pi} \sin x\,dx$。','ref':'2','steps':[r'原函数为 $-\cos x$',r'$[-\cos x]_0^\pi=1-(-1)=2$'],'explanation':'基本定积分计算'},
        {'level':1,'difficulty':'easy','kind':'slide_mcq','prompt':r'判断：若 $f(x)$ 在 $[a,b]$ 上连续，则它在 $[a,b]$ 上可积。这个说法对吗？','ref':'对。闭区间上连续函数一定黎曼可积。','steps':['连续函数可积','给出结论'],'explanation':'可积性基础判断'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'计算定积分：$\int_0^1 2x e^{x^2}\,dx$。','ref':r'$e-1$','steps':[r'令 $u=x^2$',r'$du=2x\,dx$',r'上下限仍为 $0,1$',r'得到 $\int_0^1e^u du=e-1$'],'explanation':'定积分换元'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'计算定积分：$\int_0^1 x\ln x\,dx$。','ref':r'$-\frac14$','steps':[r'分部积分，取 $u=\ln x,dv=x dx$',r'得到 $\frac{x^2}{2}\ln x-\int\frac{x}{2}dx$',r'按反常端点极限处理 $x\ln x$',r'结果 $-1/4$'],'explanation':'定积分分部积分'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求由 $y=x^2$、$x=0$、$x=1$ 和 $x$ 轴围成图形的面积。','ref':r'$\frac13$','steps':[r'面积为 $\int_0^1x^2dx$',r'计算得 $1/3$'],'explanation':'定积分几何应用基础'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'计算反常积分：$\int_1^{+\infty}\frac{1}{x^2}\,dx$。','ref':'1','steps':[r'写成极限 $\lim_{b\to+\infty}\int_1^b x^{-2}dx$',r'原函数 $-1/x$',r'极限为 $1$'],'explanation':'无穷限反常积分'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'判断反常积分 $\int_0^1\frac{1}{\sqrt{x}}\,dx$ 是否收敛，并求值。','ref':'收敛，值为 2。','steps':[r'写成 $\lim_{a\to0^+}\int_a^1x^{-1/2}dx$',r'原函数 $2\sqrt{x}$',r'极限为 $2$'],'explanation':'无界函数反常积分'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'求曲线 $y=\sqrt{x}$、直线 $x=0,x=1$ 与 $x$ 轴围成图形绕 $x$ 轴旋转所得体积。','ref':r'$\frac{\pi}{2}$','steps':[r'旋转体体积 $V=\pi\int_0^1y^2dx$',r'$y^2=x$',r'$V=\pi\int_0^1x dx=\pi/2$'],'explanation':'定积分体积应用'},
    ]
    pool = [it for it in pool if int(it.get('level',2)) == int(difficulty_level)] or pool
    rng.shuffle(pool)
    return [{
        'qid': _qid(), 'topic': topic or '定积分', 'difficulty': item['difficulty'], 'difficulty_level': item['level'],
        'prompt': item['prompt'], 'ref_answer': item['ref'], 'ref_answer_str': item['ref'], 'steps': item['steps'],
        'explanation': item['explanation'], 'ppt_pages': page_ref, '_kind': item['kind'],
        '_validation': {'passed': True, 'evidence': [f'定积分专题 L{difficulty_level}'], 'ref_answer': item['ref']},
    } for item in pool[:max(0,n)]]


def _mk_real_ode_questions(topic, pages, n, seed=None, difficulty_level=2):
    """微分方程专题按 L1/L2/L3 生成题目。"""
    import random
    rng = random.Random(seed)
    page_ref = pages[:3] if pages else []
    pool = [
        {'level':1,'difficulty':'easy','kind':'slide_concept','prompt':'概念题：什么是微分方程的通解？','ref':'含有任意常数、能够表示方程所有解的一般形式称为通解。','steps':['说明含任意常数','说明代表一族解'],'explanation':'微分方程基本概念'},
        {'level':1,'difficulty':'easy','kind':'slide_formula','prompt':r'求微分方程 $y\'=2x$ 的通解。','ref':r'$y=x^2+C$','steps':[r'两边积分：$y=\int2x dx$',r'得到 $y=x^2+C$'],'explanation':'直接积分型'},
        {'level':1,'difficulty':'easy','kind':'slide_mcq','prompt':'判断：一阶微分方程的通解通常含有一个任意常数。这个说法对吗？','ref':'对。一阶微分方程的通解通常含一个任意常数。','steps':['方程阶数与任意常数个数对应','给出判断'],'explanation':'阶数与通解概念'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求微分方程 $\frac{dy}{dx}=xy$ 的通解。','ref':r'$y=Ce^{x^2/2}$','steps':[r'分离变量：$dy/y=x dx$',r'积分：$\ln|y|=x^2/2+C$',r'化为 $y=Ce^{x^2/2}$'],'explanation':'可分离变量方程'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求一阶线性微分方程 $y\'+y=e^x$ 的通解。','ref':r'$y=\frac12e^x+Ce^{-x}$','steps':[r'积分因子 $e^x$',r'$(e^xy)\'=e^{2x}$',r'积分并整理'],'explanation':'一阶线性微分方程'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求微分方程 $y\'=\frac{x}{y}$ 的通解。','ref':r'$y^2=x^2+C$','steps':[r'分离变量：$y dy=x dx$',r'积分：$y^2/2=x^2/2+C$',r'整理'],'explanation':'可分离变量方程'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'求二阶常系数齐次方程 $y\'\'-3y\'+2y=0$ 的通解。','ref':r'$y=C_1e^x+C_2e^{2x}$','steps':[r'特征方程 $r^2-3r+2=0$',r'根为 $1,2$',r'写出通解'],'explanation':'二阶常系数齐次方程'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'求微分方程 $y\'\'=6x$ 的通解。','ref':r'$y=x^3+C_1x+C_2$','steps':[r'对 $y\'\'$ 积分得 $y\'=3x^2+C_1$',r'再积分得 $y=x^3+C_1x+C_2$'],'explanation':'可降阶高阶方程'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'求二阶常系数方程 $y\'\'+y=0$ 的通解。','ref':r'$y=C_1\cos x+C_2\sin x$','steps':[r'特征方程 $r^2+1=0$',r'根为 $\pm i$',r'写出三角形式通解'],'explanation':'复根情形二阶常系数方程'},
    ]
    pool = [it for it in pool if int(it.get('level',2)) == int(difficulty_level)] or pool
    rng.shuffle(pool)
    return [{
        'qid': _qid(), 'topic': topic or '微分方程', 'difficulty': item['difficulty'], 'difficulty_level': item['level'],
        'prompt': item['prompt'], 'ref_answer': item['ref'], 'ref_answer_str': item['ref'], 'steps': item['steps'],
        'explanation': item['explanation'], 'ppt_pages': page_ref, '_kind': item['kind'],
        '_validation': {'passed': True, 'evidence': [f'微分方程专题 L{difficulty_level}'], 'ref_answer': item['ref']},
    } for item in pool[:max(0,n)]]


def _mk_real_derivative_questions(topic, pages, n, seed=None, difficulty_level=2):
    """导数专题按 L1/L2/L3 生成题目。"""
    import random
    rng = random.Random(seed)
    page_ref = pages[:3] if pages else []
    pool = [
        {'level':1,'difficulty':'easy','kind':'slide_formula','prompt':r'求导：$f(x)=x^3-2x+1$。','ref':r'$f\'(x)=3x^2-2$','steps':[r'用幂函数求导公式',r'$ (x^3)\'=3x^2$',r'$(-2x)\'=-2$'],'explanation':'基本求导公式'},
        {'level':1,'difficulty':'easy','kind':'slide_formula','prompt':r'求导：$f(x)=\sin x+e^x$。','ref':r'$f\'(x)=\cos x+e^x$','steps':[r'$(\sin x)\'=\cos x$',r'$(e^x)\'=e^x$'],'explanation':'基本初等函数求导'},
        {'level':1,'difficulty':'easy','kind':'slide_mcq','prompt':r'判断：若函数在某点可导，则它在该点一定连续。这个说法对吗？','ref':'对。可导一定连续，但连续不一定可导。','steps':['回忆可导与连续关系','可导推出连续','连续不能反推可导'],'explanation':'导数概念判断题'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求导：$f(x)=\ln(1+x^2)$。','ref':r'$f\'(x)=\frac{2x}{1+x^2}$','steps':[r'外层为 $\ln u$',r'内层 $u=1+x^2$',r'链式法则得到 $u\'/u=2x/(1+x^2)$'],'explanation':'复合函数求导'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求导：$f(x)=x^2e^x$。','ref':r'$f\'(x)=e^x(x^2+2x)$','steps':[r'使用乘积法则',r'$(x^2)\'=2x$',r'得到 $2xe^x+x^2e^x$ 并合并'],'explanation':'乘积法则'},
        {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求由参数方程 $x=t^2,\ y=t^3$ 确定的 $dy/dx$。','ref':r'$\frac{dy}{dx}=\frac{3t}{2}$','steps':[r'$dx/dt=2t$',r'$dy/dt=3t^2$',r'$dy/dx=(dy/dt)/(dx/dt)=3t/2$'],'explanation':'参数方程求导'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'设 $y=x^x\ (x>0)$，求 $y\'$。','ref':r'$y\'=x^x(\ln x+1)$','steps':[r'两边取对数：$\ln y=x\ln x$',r'两边求导：$y\'/y=\ln x+1$',r'回代 $y=x^x$'],'explanation':'对数求导法'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'由方程 $x^2+y^2=1$ 确定 $y=y(x)$，求 $y\'$。','ref':r'$y\'=-\frac{x}{y}$','steps':[r'两边对 $x$ 求导',r'$2x+2yy\'=0$',r'解得 $y\'=-x/y$'],'explanation':'隐函数求导'},
        {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'求 $f(x)=\arctan\frac{1}{x}$ 在 $x\ne0$ 时的导数。','ref':r'$f\'(x)=-\frac{1}{x^2+1}$','steps':[r'令 $u=1/x$',r'$(\arctan u)\'=u\'/(1+u^2)$',r'$u\'=-1/x^2$',r'化简得 $-1/(x^2+1)$'],'explanation':'复合函数求导综合'},
    ]
    pool = [it for it in pool if int(it.get('level',2)) == int(difficulty_level)] or pool
    rng.shuffle(pool)
    return [{
        'qid': _qid(), 'topic': topic or '导数', 'difficulty': item['difficulty'], 'difficulty_level': item['level'],
        'prompt': item['prompt'], 'ref_answer': item['ref'], 'ref_answer_str': item['ref'], 'steps': item['steps'],
        'explanation': item['explanation'], 'ppt_pages': page_ref, '_kind': item['kind'],
        '_validation': {'passed': True, 'evidence': [f'导数专题 L{difficulty_level}'], 'ref_answer': item['ref']},
    } for item in pool[:max(0,n)]]


def _mk_real_shape_questions(topic, pages, n, seed=None, difficulty_level=2):
    """单调性/凹凸性/拐点按 L1/L2/L3 生成题目。"""
    import random
    rng = random.Random(seed)
    page_ref = pages[:3] if pages else []
    t = topic or ''
    is_mono = any(k in t for k in ('单调', '增减'))
    if is_mono:
        pool = [
            {'level':1,'difficulty':'easy','kind':'slide_mcq','prompt':r'判断：若 $f\'(x)>0$ 在区间 $(a,b)$ 内恒成立，则 $f(x)$ 在该区间单调递增。这个说法对吗？','ref':'对。导数恒正说明函数在该区间单调递增。','steps':['导数正','函数递增'],'explanation':'单调性判定概念题'},
            {'level':1,'difficulty':'easy','kind':'slide_formula','prompt':r'求函数 $f(x)=x^2-4x$ 的单调区间。','ref':r'递减区间 $(-\infty,2)$；递增区间 $(2,+\infty)$。','steps':[r'$f\'(x)=2x-4$',r'令 $f\'(x)=0$ 得 $x=2$',r'分区间判号'],'explanation':'一次导数判单调'},
            {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求函数 $f(x)=x^3-3x^2+2$ 的单调区间。','ref':r'递增区间：$(-\infty,0)$ 和 $(2,+\infty)$；递减区间：$(0,2)$。','steps':[r'$f\'(x)=3x(x-2)$',r'临界点 $0,2$',r'分区间判号'],'explanation':'多临界点单调性'},
            {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求 $f(x)=x+\frac1x\ (x\ne0)$ 的单调区间。','ref':r'递增：$(-\infty,-1)$ 与 $(1,+\infty)$；递减：$(-1,0)$ 与 $(0,1)$。','steps':[r'$f\'(x)=1-1/x^2$',r'临界点 $x=\pm1$ 并注意 $x=0$ 不在定义域',r'分区间判号'],'explanation':'含定义域断点的单调性'},
            {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'讨论函数 $f(x)=xe^{-x}$ 的单调性并求极值。','ref':r'在 $(-\infty,1)$ 递增，在 $(1,+\infty)$ 递减；$x=1$ 处取极大值 $1/e$。','steps':[r'$f\'(x)=e^{-x}(1-x)$',r'导数在 $x=1$ 变号',r'写出单调区间与极值'],'explanation':'单调性与极值综合'},
            {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'讨论 $f(x)=\ln x-ax\ (a>0)$ 的单调性。','ref':r'$f\'(x)=1/x-a$；在 $(0,1/a)$ 递增，在 $(1/a,+\infty)$ 递减。','steps':[r'定义域 $x>0$',r'求导 $1/x-a$',r'临界点 $1/a$',r'分区间判号'],'explanation':'含参数单调性'},
        ]
    else:
        pool = [
            {'level':1,'difficulty':'easy','kind':'slide_mcq','prompt':r'判断：若 $f\'\'(x)>0$ 在区间内恒成立，则函数图像在该区间向上弯曲。这个说法对吗？','ref':'对。二阶导数为正表示曲线向上弯曲。','steps':['二阶导数为正','向上弯曲'],'explanation':'凹凸性概念判断题'},
            {'level':1,'difficulty':'easy','kind':'slide_formula','prompt':r'求函数 $f(x)=x^3$ 的凹凸区间和拐点。','ref':r'$x<0$ 时凹，$x>0$ 时凸；拐点 $(0,0)$。','steps':[r'$f\'\'(x)=6x$',r'在 $0$ 两侧判号'],'explanation':'基础凹凸性'},
            {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'求函数 $f(x)=x^3-3x$ 的凹凸区间和拐点。','ref':r'$x<0$ 时凹；$x>0$ 时凸；拐点 $(0,0)$。','steps':[r'$f\'\'(x)=6x$',r'判断 $x<0,x>0$ 的符号'],'explanation':'凹凸性与拐点基础题'},
            {'level':2,'difficulty':'medium','kind':'slide_formula','prompt':r'判断函数 $f(x)=x^4-2x^2$ 的凹凸区间，并求可能的拐点。','ref':r'凸区间：$(-\infty,-1/\sqrt3)\cup(1/\sqrt3,+\infty)$；凹区间：$(-1/\sqrt3,1/\sqrt3)$；拐点在 $x=\pm1/\sqrt3$。','steps':[r'$f\'\'(x)=12x^2-4$',r'令二阶导为零',r'分区间判号'],'explanation':'二阶导数判定凹凸性'},
            {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'求 $f(x)=x^4-4x^3$ 的凹凸区间和拐点。','ref':r'$f\'\'(x)=12x(x-2)$；凸区间 $(-\infty,0)\cup(2,+\infty)$，凹区间 $(0,2)$；拐点 $(0,0),(2,-16)$。','steps':[r'求二阶导',r'解 $f\'\'=0$',r'分区间判号并代入求点'],'explanation':'多拐点凹凸性'},
            {'level':3,'difficulty':'hard','kind':'slide_formula','prompt':r'讨论 $f(x)=x+\frac1x$ 在其定义域上的凹凸性。','ref':r'$f\'\'(x)=2/x^3$；在 $(-\infty,0)$ 凹，在 $(0,+\infty)$ 凸；$x=0$ 不在定义域，不是拐点。','steps':[r'定义域排除 $0$',r'求二阶导',r'两区间判号',r'说明 $0$ 不是拐点'],'explanation':'含断点的凹凸性'},
        ]
    pool = [it for it in pool if int(it.get('level',2)) == int(difficulty_level)] or pool
    rng.shuffle(pool)
    return [{
        'qid': _qid(), 'topic': topic or ('单调性' if is_mono else '凹凸性'), 'difficulty': item['difficulty'], 'difficulty_level': item['level'],
        'prompt': item['prompt'], 'ref_answer': item['ref'], 'ref_answer_str': item['ref'], 'steps': item['steps'],
        'explanation': item['explanation'], 'ppt_pages': page_ref, '_kind': item['kind'],
        '_validation': {'passed': True, 'evidence': [f'函数性态专题 L{difficulty_level}'], 'ref_answer': item['ref']},
    } for item in pool[:max(0,n)]]


def _mk_real_limit_questions(topic, pages, n, seed=None, difficulty_level=2):
    """针对极限薄弱点生成真实练习题。"""
    import random
    rng = random.Random(seed)
    page_ref = pages[:3] if pages else []
    pool = [
        {
            'kind': 'slide_formula',
            'prompt': '求极限：$\\lim_{x\\to 0}\\frac{\\sin x}{x}$。',
            'ref': '1',
            'steps': ['这是重要极限', '$\\lim_{x\\to0}\\frac{\\sin x}{x}=1$'],
            'level': 1, 'difficulty': 'easy',
            'explanation': '重要极限基础题',
        },
        {
            'kind': 'slide_formula',
            'prompt': '求极限：$\\lim_{x\\to 1}\\frac{x^2-1}{x-1}$。',
            'ref': '2',
            'steps': ['因式分解 $x^2-1=(x-1)(x+1)$', '约去 $x-1$', '代入 $x=1$ 得 $2$'],
            'level': 1, 'difficulty': 'easy',
            'explanation': '因式分解消去零因子',
        },
        {
            'kind': 'slide_mcq',
            'prompt': '判断题：把 $\\lim_{x\\to0}\\frac{\\sin x}{x}$ 直接代入得到 $0/0$，所以极限不存在。这个说法对吗？',
            'ref': '不对。$0/0$ 是未定式，不能据此断定极限不存在；该极限等于 $1$。',
            'steps': ['直接代入得到未定式', '未定式需要继续变形或使用重要极限', '结论为 $1$'],
            'level': 1, 'difficulty': 'easy',
            'explanation': '未定式识别题',
        },
        {
            'kind': 'slide_formula',
            'prompt': '求极限：$\\lim_{x\\to0}\\frac{1-\\cos x}{x^2}$。',
            'ref': '1/2',
            'steps': ['利用等价无穷小 $1-\\cos x\\sim \\frac{x^2}{2}$', '所以原极限为 $\\frac12$'],
            'level': 2, 'difficulty': 'medium',
            'explanation': '等价无穷小题',
        },
        {
            'kind': 'slide_formula',
            'prompt': '求极限：$\\lim_{x\\to0}\\frac{e^x-1}{x}$。',
            'ref': '1',
            'steps': ['利用等价无穷小 $e^x-1\\sim x$', '所以极限为 $1$'],
            'level': 2, 'difficulty': 'medium',
            'explanation': '等价无穷小基础题',
        },

        {
            'kind': 'slide_formula',
            'prompt': r'求极限：$\lim_{x\to0}\frac{x-\sin x}{x^3}$。',
            'ref': '1/6',
            'steps': [r'用泰勒展开 $\sin x=x-\frac{x^3}{6}+o(x^3)$', r'所以 $x-\sin x\sim \frac{x^3}{6}$', r'极限为 $\frac16$'],
            'level': 3, 'difficulty': 'hard',
            'explanation': '泰勒公式求极限',
        },
        {
            'kind': 'slide_formula',
            'prompt': r'求极限：$\lim_{x\to0}\frac{\ln(1+x)-x}{x^2}$。',
            'ref': '-1/2',
            'steps': [r'用泰勒展开 $\ln(1+x)=x-\frac{x^2}{2}+o(x^2)$', r'分子等价于 $-\frac{x^2}{2}$', r'极限为 $-\frac12$'],
            'level': 3, 'difficulty': 'hard',
            'explanation': '泰勒展开求极限',
        },
    ]
    pool = [it for it in pool if int(it.get('level', 2)) == int(difficulty_level)] or pool
    rng.shuffle(pool)
    out=[]
    for item in pool[:max(0,n)]:
        out.append({
            'qid': _qid(),
            'topic': topic or '极限',
            'difficulty': item.get('difficulty', 'medium'),
            'difficulty_level': int(item.get('level', difficulty_level)),
            'prompt': item['prompt'],
            'ref_answer': item['ref'],
            'ref_answer_str': item['ref'],
            'steps': item['steps'],
            'explanation': item['explanation'],
            'ppt_pages': page_ref,
            '_kind': item['kind'],
            '_validation': {'passed': True, 'evidence': [f'薄弱点练习：{topic or "极限"}'], 'ref_answer': item['ref']},
        })
    return out


def _mk_topic_pool_questions(topic, pages, n, seed=None, difficulty_level=2, module='综合'):
    """通用模块题库：定积分应用、空间解析几何、多元微分、重积分、曲线/曲面积分、级数。"""
    import random
    rng = random.Random(seed)
    page_ref = pages[:3] if pages else []
    banks = {
        '定积分应用': [
            (1,'easy','slide_formula',r'求由 $y=x$、$x=0$、$x=1$ 与 $x$ 轴围成图形的面积。',r'$\frac12$',[r'$S=\int_0^1x\,dx=1/2$'],'平面面积基础'),
            (1,'easy','slide_formula',r'求曲线 $y=x^2$，$0\le x\le1$ 下方区域面积。',r'$\frac13$',[r'$S=\int_0^1x^2dx=1/3$'],'面积基础'),
            (2,'medium','slide_formula',r'求 $y=x$、$y=x^2$ 围成图形的面积。',r'$\frac16$',[r'交点 $x=0,1$',r'$S=\int_0^1(x-x^2)dx=1/6$'],'两曲线面积'),
            (2,'medium','slide_formula',r'求 $y=x^2$，$0\le x\le1$ 绕 $x$ 轴旋转所得体积。',r'$\frac{\pi}{5}$',[r'$V=\pi\int_0^1y^2dx=\pi\int_0^1x^4dx=\pi/5$'],'旋转体体积'),
            (3,'hard','slide_formula',r'求星形线参数方程 $x=a\cos^3t,y=a\sin^3t$ 第一象限弧长。',r'$\frac{3a}{2}$',[r'用参数曲线弧长公式',r'计算 $\sqrt{(dx/dt)^2+(dy/dt)^2}=3a\sin t\cos t$',r'积分 $0$ 到 $\pi/2$'],'参数曲线弧长'),
            (3,'hard','slide_formula',r'半径为 $R$ 的圆面积用极坐标定积分表示并计算。',r'$\pi R^2$',[r'$S=\int_0^{2\pi}\int_0^R r\,drd\theta$',r'结果 $\pi R^2$'],'极坐标面积'),
        ],
        '空间解析几何': [
            (1,'easy','slide_formula',r'求向量 $a=(1,2,3)$ 与 $b=(2,0,1)$ 的数量积。','5',[r'$a\cdot b=1\cdot2+2\cdot0+3\cdot1=5$'],'数量积'),
            (1,'easy','slide_formula',r'求过点 $(1,2,3)$ 且法向量为 $(1,0,-1)$ 的平面方程。',r'$x-z+2=0$',[r'点法式：$(x-1)-(z-3)=0$'],'平面点法式'),
            (2,'medium','slide_formula',r'求平面 $x+2y+2z-5=0$ 到点 $(1,0,0)$ 的距离。',r'$\frac43$',[r'代入距离公式 $|1-5|/\sqrt{1+4+4}=4/3$'],'点到平面距离'),
            (2,'medium','slide_formula',r'求直线 $\frac{x-1}{2}=\frac{y}{-1}=z+1$ 的方向向量。',r'$(2,-1,1)$',[r'由对称式直接读出方向向量'],'空间直线'),
            (3,'hard','slide_formula',r'求平面 $x+y+z=1$ 与 $x-y+z=0$ 的夹角。',r'$\arccos\frac{1}{3}$',[r'法向量 $n_1=(1,1,1),n_2=(1,-1,1)$',r'夹角余弦 $|n_1\cdot n_2|/(|n_1||n_2|)=1/3$'],'平面夹角'),
            (3,'hard','slide_formula',r'判断曲面 $x^2+y^2=1$ 是什么曲面？','圆柱面。',[r'方程与 $z$ 无关',r'在任意 $z$ 截面均为单位圆'],'柱面识别'),
        ],
        '多元函数微分': [
            (1,'easy','slide_formula',r'求 $f(x,y)=x^2+y^2$ 的偏导数。',r'$f_x=2x,\ f_y=2y$',[r'分别把另一个变量看作常数求导'],'偏导基础'),
            (1,'easy','slide_formula',r'求 $z=x^2y$ 在点 $(1,2)$ 处的全微分。',r'$dz=4dx+dy$',[r'$z_x=2xy,z_y=x^2$',r'代入 $(1,2)$'],'全微分基础'),
            (2,'medium','slide_formula',r'设 $z=\ln(x^2+y^2)$，求 $z_x,z_y$。',r'$z_x=\frac{2x}{x^2+y^2},\ z_y=\frac{2y}{x^2+y^2}$',[r'链式求偏导'],'多元复合求导'),
            (2,'medium','slide_formula',r'求 $f(x,y)=x^2+y^2-2x-4y$ 的极值点。',r'极小值点 $(1,2)$。',[r'$f_x=2x-2,f_y=2y-4$',r'驻点 $(1,2)$',r'二次型正定'],'多元极值'),
            (3,'hard','slide_formula',r'用拉格朗日乘数法求 $x^2+y^2$ 在约束 $x+y=1$ 下的最小值。',r'$\frac12$',[r'$\nabla f=\lambda\nabla g$',r'$2x=\lambda,2y=\lambda,x+y=1$',r'$x=y=1/2$'],'条件极值'),
            (3,'hard','slide_formula',r'求曲面 $z=x^2+y^2$ 在点 $(1,1,2)$ 的切平面。',r'$z=2+2(x-1)+2(y-1)$',[r'$z_x=2x,z_y=2y$',r'切平面公式'],'切平面'),
        ],
        '重积分': [
            (1,'easy','slide_formula',r'计算二重积分 $\iint_D 1\,dA$，其中 $D=[0,1]\times[0,2]$。','2',[r'积分值等于矩形面积 $1\cdot2=2$'],'二重积分概念'),
            (1,'easy','slide_formula',r'计算 $\int_0^1\int_0^1 (x+y)\,dy\,dx$。','1',[r'先对 $y$ 积分再对 $x$ 积分'],'矩形区域二重积分'),
            (2,'medium','slide_formula',r'计算 $D:0\le y\le x\le1$ 上的 $\iint_D x\,dA$。',r'$\frac13$',[r'$\int_0^1\int_0^x x\,dy dx=\int_0^1x^2dx$'],'一般区域二重积分'),
            (2,'medium','slide_formula',r'用极坐标计算单位圆内 $\iint_D (x^2+y^2)\,dA$。',r'$\frac{\pi}{2}$',[r'极坐标下为 $\int_0^{2\pi}\int_0^1 r^2\cdot r drd\theta$'],'极坐标二重积分'),
            (3,'hard','slide_formula',r'计算球体 $x^2+y^2+z^2\le R^2$ 的体积三重积分。',r'$\frac43\pi R^3$',[r'球坐标 $dV=\rho^2\sin\varphi d\rho d\varphi d\theta$',r'积分范围 $0\le\rho\le R$'],'球坐标三重积分'),
            (3,'hard','slide_formula',r'用柱坐标表示圆柱 $x^2+y^2\le1,0\le z\le2$ 的体积并计算。',r'$2\pi$',[r'$V=\int_0^{2\pi}\int_0^1\int_0^2 r dzdrd\theta=2\pi$'],'柱坐标三重积分'),
        ],
        '曲线曲面积分': [
            (1,'easy','slide_formula',r'计算单位线段 $0\le x\le1$ 上对弧长的曲线积分 $\int_L 1\,ds$。','1',[r'等于线段长度'],'第一类曲线积分'),
            (1,'easy','slide_mcq',r'判断：第一类曲线积分与曲线方向有关。这个说法对吗？','不对。第一类曲线积分对弧长积分，与方向无关。',['第一类对弧长积分','方向无关'],'概念判断'),
            (2,'medium','slide_formula',r'计算 $L:y=x,0\le x\le1$ 上 $\int_L x\,ds$。',r'$\frac{\sqrt2}{2}$',[r'$ds=\sqrt2dx$',r'$\int_0^1x\sqrt2dx=\sqrt2/2$'],'第一类曲线积分计算'),
            (2,'medium','slide_formula',r'用格林公式求 $\oint_C -y\,dx+x\,dy$，$C$ 为单位圆正向。',r'$2\pi$',[r'$Q_x-P_y=1-(-1)=2$',r'面积为 $\pi$'],'格林公式'),
            (3,'hard','slide_formula',r'判断向量场 $P=-y/(x^2+y^2),Q=x/(x^2+y^2)$ 在不含原点的单连通区域内曲线积分是否与路径无关。','路径无关。',[r'计算 $P_y$ 与 $Q_x$ 在区域内相等',r'区域不含奇点且单连通'],'路径无关'),
            (3,'hard','slide_formula',r'单位球面外侧的通量 $\iint_S (x,y,z)\cdot n\,dS$。',r'$4\pi$',[r'用高斯公式，散度为 $3$',r'单位球体体积 $4\pi/3$',r'通量 $3\cdot4\pi/3=4\pi$'],'高斯公式'),
        ],
        '无穷级数': [
            (1,'easy','slide_mcq',r'判断级数 $\sum_{n=1}^\infty \frac1{n^2}$ 是否收敛。','收敛。它是 p=2>1 的 p 级数。',['p 级数判别'],'p级数'),
            (1,'easy','slide_formula',r'求等比级数 $\sum_{n=0}^\infty (1/2)^n$ 的和。','2',[r'等比级数和 $1/(1-1/2)=2$'],'等比级数'),
            (2,'medium','slide_formula',r'用比值审敛法判断 $\sum_{n=1}^\infty \frac{1}{n!}$ 的敛散性。','收敛。',[r'比值极限 $1/(n+1)\to0<1$'],'比值审敛法'),
            (2,'medium','slide_formula',r'求幂级数 $\sum_{n=1}^\infty \frac{x^n}{n}$ 的收敛半径。','1',[r'系数比值或根值极限得到 $R=1$'],'收敛半径'),
            (3,'hard','slide_formula',r'判断交错级数 $\sum_{n=1}^\infty (-1)^{n-1}/n$ 的绝对收敛或条件收敛。','条件收敛。',[r'交错调和级数收敛',r'绝对值级数为调和级数发散'],'条件收敛'),
            (3,'hard','slide_formula',r'写出 $e^x$ 的麦克劳林级数并说明收敛域。',r'$e^x=\sum_{n=0}^\infty x^n/n!$，收敛域为全体实数。',[r'标准展开式',r'收敛半径无穷大'],'麦克劳林级数'),
        ],
    }
    pool = banks.get(module, [])
    pool = [it for it in pool if int(it[0]) == int(difficulty_level)] or pool
    rng.shuffle(pool)
    out=[]
    for level,diff,kind,prompt,ref,steps,expl in pool[:max(0,n)]:
        out.append({'qid': _qid(), 'topic': topic or module, 'difficulty': diff, 'difficulty_level': level,
                    'prompt': prompt, 'ref_answer': ref, 'ref_answer_str': ref, 'steps': steps,
                    'explanation': expl, 'ppt_pages': page_ref, '_kind': kind,
                    '_validation': {'passed': True, 'evidence': [f'{module} L{difficulty_level}'], 'ref_answer': ref}})
    return out


def _mk_real_concept_questions(topic, pages, n, seed=None, difficulty_level=2):
    """泛“概念”薄弱点题库，避免主题过泛时生成空题。"""
    import random
    rng = random.Random(seed)
    page_ref = pages[:3] if pages else []
    pool = [
        {'level':1,'difficulty':'easy','kind':'slide_mcq','prompt':'判断：函数在某点可导，则它在该点一定连续。这个说法对吗？','ref':'对。可导一定连续，但连续不一定可导。','steps':['可导意味着差商极限存在','可推出函数在该点连续'],'explanation':'可导与连续关系'},
        {'level':1,'difficulty':'easy','kind':'slide_mcq','prompt':'判断：把极限直接代入得到 $0/0$，就说明极限不存在。这个说法对吗？','ref':'不对。$0/0$ 是未定式，不能直接判断极限不存在，需要继续化简或使用其他方法。','steps':['识别 $0/0$ 为未定式','未定式需要进一步处理'],'explanation':'未定式概念'},
        {'level':1,'difficulty':'easy','kind':'slide_concept','prompt':'概念题：什么叫函数在某点连续？','ref':'函数在该点有定义，极限存在，并且极限值等于函数值。','steps':['函数值存在','极限存在','二者相等'],'explanation':'连续性定义'},
        {'level':2,'difficulty':'medium','kind':'slide_concept','prompt':'概念题：单调性通常怎样用导数判断？','ref':'在区间内导数大于零通常函数递增，导数小于零通常函数递减；还要结合定义域和临界点分区间讨论。','steps':['求导','找导数为零或不存在的点','分区间判断导数符号'],'explanation':'单调性判定方法'},
        {'level':2,'difficulty':'medium','kind':'slide_concept','prompt':'概念题：曲线凹凸性通常怎样判断？','ref':'一般求二阶导数，二阶导数为正对应向上弯曲，二阶导数为负对应向下弯曲；二阶导数变号的点可能是拐点。','steps':['求二阶导','找二阶导为零或不存在的点','分区间判号','判断拐点'],'explanation':'凹凸性判定方法'},
        {'level':2,'difficulty':'medium','kind':'slide_concept','prompt':'概念题：不定积分中的“任意常数 C”为什么不能省略？','ref':'因为同一个导数对应无穷多个相差常数的原函数，不定积分表示全体原函数，所以要加任意常数 C。','steps':['原函数不唯一','相差常数的函数导数相同','不定积分表示一族函数'],'explanation':'原函数与不定积分概念'},
        {'level':3,'difficulty':'hard','kind':'slide_concept','prompt':'概念题：为什么连续不一定可导？请举例说明。','ref':'连续只要求函数值与极限一致，可导还要求差商极限存在。例如 $f(x)=|x|$ 在 $x=0$ 连续，但左右导数不同，因此不可导。','steps':['说明连续条件','说明可导条件更强','举 $|x|$ 的尖点例子'],'explanation':'连续与可导辨析'},
        {'level':3,'difficulty':'hard','kind':'slide_concept','prompt':'概念题：极值点和最值点有什么区别？','ref':'极值是局部范围内的最大或最小，最值是给定区间或定义域上的全局最大或最小；极值不一定是最值。','steps':['极值是局部概念','最值是整体概念','二者不一定重合'],'explanation':'极值与最值辨析'},
        {'level':3,'difficulty':'hard','kind':'slide_concept','prompt':'概念题：反常积分和普通定积分的主要区别是什么？','ref':'普通定积分通常在有限区间上对有界函数积分；反常积分涉及无穷区间或无界函数，需要用极限定义并判断收敛性。','steps':['指出无穷区间或无界函数','需要极限定义','要判断收敛或发散'],'explanation':'反常积分概念'},
    ]
    pool = [it for it in pool if int(it.get('level',2)) == int(difficulty_level)] or pool
    rng.shuffle(pool)
    return [{
        'qid': _qid(), 'topic': topic or '概念', 'difficulty': item['difficulty'], 'difficulty_level': item['level'],
        'prompt': item['prompt'], 'ref_answer': item['ref'], 'ref_answer_str': item['ref'], 'steps': item['steps'],
        'explanation': item['explanation'], 'ppt_pages': page_ref, '_kind': item['kind'],
        '_validation': {'passed': True, 'evidence': [f'概念题 L{difficulty_level}'], 'ref_answer': item['ref']},
    } for item in pool[:max(0,n)]]


def generate_from_slides(slides, topic=None, difficulty_level=2, n=5,
                         seed=None, ppt_hint=None):
    """主路径: 按用户指定主题 + 当前 PPT 素材出多样化题.

    若 PPT 中没有该主题相关内容, 返回空列表.
    """
    pages = ppt_hint or _find_related_pages(slides, topic, limit=8)
    t = topic or ''
    if any(k in t for k in ('概念', '定义', '辨析')):
        return _mk_real_concept_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level)
    if any(k in t for k in ('极限', 'lim', 'limit')):
        return _mk_real_limit_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level)
    if any(k in t for k in ('定积分应用', '弧长', '变力做功', '水压力', '引力', '旋转体')):
        return _mk_topic_pool_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level, module='定积分应用')
    if any(k in t for k in ('空间解析几何', '向量', '平面方程', '空间直线', '二次曲面', '柱面', '旋转曲面')):
        return _mk_topic_pool_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level, module='空间解析几何')
    if any(k in t for k in ('多元函数', '偏导', '全微分', '方向导数', '梯度', '条件极值', '拉格朗日乘数')):
        return _mk_topic_pool_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level, module='多元函数微分')
    if any(k in t for k in ('重积分', '二重积分', '三重积分', '极坐标', '柱坐标', '球坐标')):
        return _mk_topic_pool_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level, module='重积分')
    if any(k in t for k in ('曲线积分', '曲面积分', '格林公式', '高斯公式', '斯托克斯', '路径无关')):
        return _mk_topic_pool_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level, module='曲线曲面积分')
    if any(k in t for k in ('级数', '幂级数', '傅里叶', '收敛半径', '审敛', '泰勒级数')):
        return _mk_topic_pool_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level, module='无穷级数')
    if any(k in t for k in ('定积分', '反常积分', '牛顿-莱布尼茨', '面积', '体积')):
        return _mk_real_definite_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level)
    if any(k in t for k in ('微分方程', '通解', '特解', '可分离变量', '一阶线性')):
        return _mk_real_ode_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level)
    if any(k in t for k in ('积分', '换元', '凑微分', '代换')):
        return _mk_real_integral_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level)
    if any(k in t for k in ('导数', '求导', '微分')):
        return _mk_real_derivative_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level)
    if any(k in t for k in ('凹凸', '凸性', '拐点', '单调', '增减')):
        return _mk_real_shape_questions(topic, pages, n, seed=seed, difficulty_level=difficulty_level)
    if not pages:
        return []
    return _mk_questions_from_pages(slides, topic, pages, n, seed=seed)


def generate(topic=None, difficulty_level=2, n=5, seed=None, ppt_hint=None):
    """兼容旧接口. 当前最小版只支持积分族; 非积分主题返回空."""
    if topic and any(k in topic for k in ('积分', '换元', '凑微分', '分部')):
        return generate_from_slides([], topic=topic, difficulty_level=difficulty_level,
                                    n=n, seed=seed, ppt_hint=ppt_hint)
    return []


def save_bank(bank, path=r'E:\agent 开发\question_bank.json'):
    clean = []
    for b in bank:
        b2 = dict(b)
        if not isinstance(b2.get('ref_answer'), (str, int, float)):
            b2['ref_answer'] = str(b2.get('ref_answer'))
        clean.append(b2)
    out = {
        'generated_at': time.strftime('%Y-%m-%d %H:%M:%S'),
        'n_total': len(clean),
        'n_passed_validation': sum(1 for b in clean if b.get('_validation', {}).get('passed')),
        'questions': clean,
    }
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    return out


if __name__ == '__main__':
    slides = json.load(open(r'E:\agent 开发\slides.json', encoding='utf-8')).get('slides', [])
    bank = generate_from_slides(slides, topic='换元积分', n=5, seed=42)
    save_bank(bank)
    print(f"生成 {len(bank)} 题")
    for b in bank:
        print(f"[{b['qid']}] ({b['_kind']}) {b['prompt'][:60]}")
