#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
西班牙语动词变位 —— 听觉主导学习工具
=============================================
功能：
  1. 浏览模式 — 逐个动词浏览，听+看 6 种人称变位（不规则高亮+解释）
  2. 测验模式 — 听读音，回答动词原形+主语人称(可选中文)，计分
  3. 听写模式 — 听原形+人称，写出变位形式，计分
  4. 设置 — 语速/音量/语音

依赖：pywin32, colorama
安装：pip install pywin32 colorama
"""

import sys
import os
import time
import random
import traceback
import json

# 修复 Windows GBK 终端无法编码西班牙语字符（¡¿áéíóúñ）的问题
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_error_log.txt')
STARRED_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_starred.json')

# ---------- 星标动词管理 ----------
starred_verbs = set()  # 存储星标动词的原形

def load_starred():
    global starred_verbs
    try:
        with open(STARRED_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            starred_verbs = set(data.get('verbs', []))
    except (FileNotFoundError, json.JSONDecodeError):
        starred_verbs = set()

def save_starred():
    try:
        with open(STARRED_FILE, 'w', encoding='utf-8') as f:
            json.dump({'verbs': list(starred_verbs)}, f, ensure_ascii=False)
    except:
        pass

def log_error(msg: str):
    try:
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write(msg)
    except:
        pass

# ---------- colorama ----------
try:
    from colorama import init, Fore, Back, Style
    init(autoreset=True)
except ImportError:
    class _Fake:
        def __getattr__(self, name):
            return ''
    Fore = _Fake()
    Back = _Fake()
    Style = _Fake()
    def init(*a, **kw): pass

# ---------- SAPI TTS ----------
try:
    import win32com.client
    HAS_TTS = True
except ImportError:
    HAS_TTS = False

# ============================================================
#  动词数据
# ============================================================

SUBJECTS = [
    ("yo",               "我"),
    ("tú",               "你"),
    ("él / ella / usted",    "他/她/您"),
    ("nosotros / nosotras","我们"),
    ("vosotros / vosotras","你们"),
    ("ellos / ellas / ustedes","他们/她们/诸位"),
]

# 朗读时把复合主语拆成独立条目，每个单独连读
# 每个元素: (屏幕显示文本, [实际朗读的独立西语主语列表])
SUBJECT_SPLIT = [
    ("yo",                       ["yo"]),
    ("tú",                       ["tú"]),
    ("él / ella / usted",        ["él", "ella", "usted"]),
    ("nosotros / nosotras",      ["nosotros", "nosotras"]),
    ("vosotros / vosotras",      ["vosotros", "vosotras"]),
    ("ellos / ellas / ustedes",  ["ellos", "ellas", "ustedes"]),
]

# 与 SUBJECT_SPLIT 一一对应的中文主语，按子项拆分
# e.g. él→他  ella→她  usted→您  (而非"他/她/您"混在一起)
SUBJECTS_CN_SPLIT = [
    ["我"],
    ["你"],
    ["他", "她", "您"],
    ["我们", "我们"],
    ["你们", "你们"],
    ["他们", "她们", "诸位"],
]

def pick_random_subject(idx: int):
    """从复合主语中随机选一个，返回 (西语主语, 中文主语)"""
    parts_es = SUBJECT_SPLIT[idx][1]
    parts_cn = SUBJECTS_CN_SPLIT[idx]
    i = random.randint(0, len(parts_es) - 1)
    return parts_es[i], parts_cn[i]

VERBS = [
    # ========== 第一组：规则 -ar 动词 ==========
    {
        "infinitive": "hablar",
        "cn": "说话",
        "category": "第一组 规则 -ar 动词",
        "conjugations": ["hablo", "hablas", "habla", "hablamos", "habláis", "hablan"],
        "example": "",
    },
    {
        "infinitive": "cantar",
        "cn": "唱",
        "category": "第一组 规则 -ar 动词",
        "conjugations": ["canto", "cantas", "canta", "cantamos", "cantáis", "cantan"],
        "example": "",
    },
    {
        "infinitive": "estudiar",
        "cn": "学习",
        "category": "第一组 规则 -ar 动词",
        "conjugations": ["estudio", "estudias", "estudia", "estudiamos", "estudiáis", "estudian"],
        "example": "Estudio mucho. -> 我学习很努力。\n¿Qué estudias? -> 你是学什么的？\nEstudio español. -> 我是学西班牙语的。",
    },
    # ========== 第二组：规则 -er 动词 ==========
    {
        "infinitive": "comer",
        "cn": "吃",
        "category": "第二组 规则 -er 动词",
        "conjugations": ["como", "comes", "come", "comemos", "coméis", "comen"],
        "example": "",
    },
    {
        "infinitive": "beber",
        "cn": "喝",
        "category": "第二组 规则 -er 动词",
        "conjugations": ["bebo", "bebes", "bebe", "bebemos", "bebéis", "beben"],
        "example": "",
    },
    # ========== 第三组：规则 -ir 动词 ==========
    {
        "infinitive": "cumplir",
        "cn": "完成",
        "category": "第三组 规则 -ir 动词",
        "conjugations": ["cumplo", "cumples", "cumple", "cumplimos", "cumplís", "cumplen"],
        "example": "",
    },
    {
        "infinitive": "subir",
        "cn": "上升",
        "category": "第三组 规则 -ir 动词",
        "conjugations": ["subo", "subes", "sube", "subimos", "subís", "suben"],
        "example": "",
    },
    # ========== 不规则动词 ==========
    {
        "infinitive": "ser",
        "cn": "是",
        "category": "不规则动词",
        "conjugations": ["soy", "eres", "es", "somos", "sois", "son"],
        "example": "",
        "highlights": {
            0: ("soy", "完全换词根"),
            1: ("eres", "完全换词根"),
            2: ("es",   "完全换词根"),
            3: ("somos", "完全换词根"),
            4: ("sois",  "完全换词根"),
            5: ("son",   "完全换词根"),
        },
    },
    {
        "infinitive": "estar",
        "cn": "处于、处在",
        "category": "不规则动词",
        "conjugations": ["estoy", "estás", "está", "estamos", "estáis", "están"],
        "example": "",
        "highlights": {
            0: ("oy", "yo 形 +y 缓冲"),
            1: ("ás", "重音后移 → ás"),
            2: ("á",  "重音后移 → á"),
            5: ("án", "重音后移 → án"),
        },
    },
    {
        "infinitive": "tener",
        "cn": "有",
        "category": "不规则动词",
        "conjugations": ["tengo", "tienes", "tiene", "tenemos", "tenéis", "tienen"],
        "example": "",
        "highlights": {
            0: ("go", "yo 形 +go 缓冲辅音"),
            1: ("ie", "e→ie 元音裂化"),
            2: ("ie", "e→ie 元音裂化"),
            5: ("ie", "e→ie 元音裂化"),
        },
    },
    {
        "infinitive": "poder",
        "cn": "能够",
        "category": "不规则动词",
        "conjugations": ["puedo", "puedes", "puede", "podemos", "podéis", "pueden"],
        "example": "",
        "highlights": {
            0: ("ue", "o→ue 元音裂化"),
            1: ("ue", "o→ue 元音裂化"),
            2: ("ue", "o→ue 元音裂化"),
            5: ("ue", "o→ue 元音裂化"),
        },
    },
    {
        "infinitive": "poner",
        "cn": "放",
        "category": "不规则动词",
        "conjugations": ["pongo", "pones", "pone", "ponemos", "ponéis", "ponen"],
        "example": "",
        "highlights": {
            0: ("go", "yo 形 +go 缓冲辅音"),
        },
    },
    {
        "infinitive": "venir",
        "cn": "来",
        "category": "不规则动词",
        "conjugations": ["vengo", "vienes", "viene", "venimos", "venís", "vienen"],
        "example": "",
        "highlights": {
            0: ("go", "yo 形 +go 缓冲辅音（同 tener）"),
            1: ("ie", "e→ie 元音裂化"),
            2: ("ie", "e→ie 元音裂化"),
            5: ("ie", "e→ie 元音裂化"),
        },
    },
    {
        "infinitive": "ir",
        "cn": "去",
        "category": "不规则动词",
        "conjugations": ["voy", "vas", "va", "vamos", "vais", "van"],
        "example": "",
        "highlights": {
            0: ("voy", "完全换词根"),
            1: ("vas", "完全换词根"),
            2: ("va",  "完全换词根"),
            3: ("vamos", "完全换词根"),
            4: ("vais",  "完全换词根"),
            5: ("van",   "完全换词根"),
        },
    },
]

# ============================================================
#  SAPI TTS 引擎
# ============================================================

class Speaker:
    """Windows SAPI 语音引擎 —— 逐句稳定朗读"""

    def __init__(self):
        self.tts_enabled = False
        self.voice = None
        self.voice_es = None  # 西班牙语语音对象
        self.voice_cn = None  # 中文语音对象
        self.rate = 0
        self.volume = 100

        if not HAS_TTS:
            print(Fore.RED + "  [警告] pywin32 未安装，语音功能禁用。pip install pywin32")
            return
        try:
            self.voice = win32com.client.Dispatch("SAPI.SpVoice")
            voices = self.voice.GetVoices()
            for v in voices:
                name = v.GetDescription().lower()
                if not self.voice_es and ('spanish' in name or 'español' in name):
                    self.voice_es = v
                if not self.voice_cn and ('chinese' in name or 'huihui' in name):
                    self.voice_cn = v
            # 回退：找不到就用列表里第一个英语当西语，第一个中文当中文
            if not self.voice_es and len(voices) > 0:
                self.voice_es = voices[0]
            if not self.voice_cn and len(voices) > 1:
                self.voice_cn = voices[1]  # 一般第二个是中文
            elif not self.voice_cn:
                self.voice_cn = self.voice_es

            self.voice.Voice = self.voice_es
            self.voice.Rate = 0
            self.voice.Volume = 100
            self.tts_enabled = True
        except Exception as e:
            print(Fore.RED + f"  [警告] TTS 引擎初始化失败: {e}")
            self.tts_enabled = False

    def _is_chinese(self, text: str) -> bool:
        """检测文本是否包含中文"""
        for ch in text:
            if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f':
                return True
        return False

    def say(self, text: str):
        """朗读文本 —— 自动检测中/西并用对应语音"""
        if not self.tts_enabled or not text.strip():
            return
        try:
            if self._is_chinese(text) and self.voice_cn:
                self.voice.Voice = self.voice_cn
            elif not self._is_chinese(text) and self.voice_es:
                self.voice.Voice = self.voice_es
            self.voice.Speak(text)
        except Exception:
            pass

    def set_rate(self, rate: int):
        """语速 -10 ~ 10"""
        if self.tts_enabled:
            self.voice.Rate = rate

    def set_volume(self, vol: int):
        """音量 0 ~ 100"""
        if self.tts_enabled:
            self.voice.Volume = vol

    def get_rate(self):
        return self.voice.Rate if self.tts_enabled else "N/A"

    def get_volume(self):
        return self.voice.Volume if self.tts_enabled else "N/A"

    def list_voices(self):
        if self.tts_enabled:
            voices = self.voice.GetVoices()
            current = self.voice.Voice.GetDescription()
            for i, v in enumerate(voices):
                marker = " <-- 当前" if v.GetDescription() == current else ""
                print(f"  [{i}] {v.GetDescription()}{marker}")
            return voices
        else:
            print("  (语音引擎不可用)")
            return []

    def set_voice(self, index: int):
        if self.tts_enabled:
            voices = self.voice.GetVoices()
            if 0 <= index < len(voices):
                self.voice.Voice = voices[index]

# ============================================================
#  打印工具
# ============================================================

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def print_title(text: str):
    print()
    print(Fore.CYAN + Style.BRIGHT + "=" * 60)
    print(Fore.YELLOW + Style.BRIGHT + f"  {text}")
    print(Fore.CYAN + Style.BRIGHT + "=" * 60)
    print()

def print_verb_header(verb: dict, idx: int, total: int):
    star_marker = " ⭐" if verb['infinitive'] in starred_verbs else ""
    print(Fore.MAGENTA + Style.BRIGHT + f"--- [{idx+1}/{total}]{star_marker} ---")
    print(Fore.GREEN + Style.BRIGHT + f"  动词原形: {verb['infinitive']}")
    print(Fore.WHITE + f"  中文意思: {verb['cn']}")
    print(Fore.CYAN + f"  所属类别: {verb['category']}")
    if verb.get('example'):
        print(Fore.YELLOW + f"  例句: {verb['example']}")
    print(Fore.MAGENTA + Style.BRIGHT + "-" * 40)

def make_cn_phrase(person_idx: int, sub_idx: int, verb_cn: str) -> str:
    """构建「我说话」「他说话」这样的中文短语（单主语+动词）"""
    subj_cn = SUBJECTS_CN_SPLIT[person_idx][sub_idx]
    return f"{subj_cn}{verb_cn}"

def visible_len(s: str) -> int:
    """计算去掉 ANSI 转义码后的可见字符长度"""
    import re
    return len(re.sub(r'\x1b\[[0-9;]*m', '', s))

def pad_visible(s: str, width: int) -> str:
    """在字符串后补空格到可见宽度 width"""
    need = width - visible_len(s)
    return s + (' ' * max(0, need))

def format_highlighted_conj(conj: str, hl_info, has_irregular: bool) -> str:
    """给变位着色：不规则字母红色高亮，其余黄色。规则动词为亮黄，不规则动词中规则位置为亮绿"""
    if hl_info:
        hl_letters, _ = hl_info
        idx_pos = conj.lower().find(hl_letters.lower())
        if idx_pos >= 0:
            before = conj[:idx_pos]
            hl_part = conj[idx_pos:idx_pos + len(hl_letters)]
            after = conj[idx_pos + len(hl_letters):]
            return (Fore.YELLOW + Style.BRIGHT + before +
                    Fore.RED + Style.BRIGHT + hl_part +
                    Fore.YELLOW + Style.BRIGHT + after)
        else:
            return Fore.RED + Style.BRIGHT + conj
    else:
        if has_irregular:
            return Fore.GREEN + Style.BRIGHT + conj
        else:
            return Fore.YELLOW + Style.BRIGHT + conj


def print_conjugation_table(verb: dict, highlight_idx: int = -1):
    """打印变位表。一行一个变位，不规则字母红色高亮，解释紧跟行尾。"""
    highlights = verb.get('highlights', {})
    has_irregular = bool(highlights)

    print()
    if has_irregular:
        print(Fore.MAGENTA + Style.BRIGHT + "  🔴 红色字母 = 不规则变化  |  "
              + Fore.GREEN + "nosotros/vosotros 通常规则")
        print()

    print(Fore.WHITE + Style.BRIGHT + f"  {'主格人称代词':<26}{'变位':<14}  {'西 → 中'}")
    print(Fore.WHITE + "  " + "-" * 62)
    for i, (subj_es, subj_cn) in enumerate(SUBJECTS):
        conj = verb['conjugations'][i]
        cn_phrase = make_cn_phrase(i, 0, verb['cn'])
        marker = Fore.CYAN + " ◀" if i == highlight_idx else ""

        # 变位着色
        if i in highlights:
            hl_letters, hl_note = highlights[i]
            conj_col = format_highlighted_conj(conj, highlights[i], has_irregular)
            conj_plain = conj
            note = f"  {Fore.RED}{hl_note}"
        else:
            conj_col = Fore.GREEN + Style.BRIGHT + conj if has_irregular else Fore.YELLOW + Style.BRIGHT + conj
            conj_plain = conj
            note = ""

        print(f"  {Fore.CYAN + subj_es:<26}{pad_visible(conj_col, 16)} {Fore.WHITE}{conj_plain:<12} → {Fore.WHITE}{cn_phrase}{marker}{note}")

    print()

def print_divider():
    print(Fore.BLUE + "-" * 60)

def speak_conjugations(speaker: Speaker, verb: dict):
    """朗读全部变位 —— 西语 + 中文都念，每个子主语配匹配的中文"""
    highlights = verb.get('highlights', {})
    has_irregular = bool(highlights)
    print(Fore.GREEN + ">> 朗读全部变位（西语 → 中文，复合主语已拆开）...")
    for i, (subj_display, _) in enumerate(SUBJECT_SPLIT):
        conj = verb['conjugations'][i]
        conj_col = format_highlighted_conj(conj, highlights.get(i), has_irregular)
        parts_es = SUBJECT_SPLIT[i][1]
        parts_cn = SUBJECTS_CN_SPLIT[i]
        for j, subj_es in enumerate(parts_es):
            cn_phrase = make_cn_phrase(i, j, verb['cn'])
            es_phrase = f"{subj_es} {conj}"
            print(f"     {Fore.CYAN + subj_es:<14} {pad_visible(conj_col, 16)} → {Fore.GREEN + cn_phrase}")
            speaker.say(es_phrase)
            time.sleep(0.15)
            speaker.say(cn_phrase)
            time.sleep(0.2)
        time.sleep(0.1)
    print(Fore.GREEN + "  [OK] 朗读完毕！")

def speak_conjugations_shuffle(speaker: Speaker, verb: dict, mode: str):
    """乱序朗读全部变位
    mode: 'cn'=西语+中文  'fast'=仅西语连读  'pause'=仅西语主语停顿
          'cnfirst'=中文先出(等反应)再西语"""
    highlights = verb.get('highlights', {})
    has_irregular = bool(highlights)
    all_pairs = []
    for i in range(6):
        conj = verb['conjugations'][i]
        conj_col = format_highlighted_conj(conj, highlights.get(i), has_irregular)
        parts_es = SUBJECT_SPLIT[i][1]
        parts_cn = SUBJECTS_CN_SPLIT[i]
        for j, subj_es in enumerate(parts_es):
            cn_subj = parts_cn[j]
            cn_phrase = f"{cn_subj}{verb['cn']}"
            all_pairs.append((subj_es, conj, conj_col, cn_phrase))
    random.shuffle(all_pairs)

    mode_desc = {
        "cn": "西语 + 中文",
        "fast": "仅西语（连读）",
        "pause": "仅西语（主语-停顿-变位）",
        "cnfirst": "中文先出（等反应）→ 西语验证"
    }[mode]
    print(Fore.GREEN + f">> 乱序朗读全部变位（{mode_desc}）...")
    for subj_es, conj, conj_col, cn_phrase in all_pairs:
        if mode == 'cn':
            print(f"     {Fore.CYAN + subj_es:<14} {pad_visible(conj_col, 16)} → {Fore.GREEN + cn_phrase}")
            speaker.say(f"{subj_es} {conj}")
            time.sleep(0.15)
            speaker.say(cn_phrase)
            time.sleep(0.2)
        elif mode == 'fast':
            print(f"     {Fore.CYAN + subj_es:<14} {conj_col}")
            speaker.say(f"{subj_es} {conj}")
            time.sleep(0.12)
        elif mode == 'pause':
            print(f"     {Fore.CYAN + subj_es:<14} {conj_col}")
            speaker.say(subj_es)
            time.sleep(0.35)
            speaker.say(conj)
            time.sleep(0.15)
        elif mode == 'cnfirst':
            print(f"     {Fore.GREEN + cn_phrase:<12} → {Fore.CYAN + subj_es:<14} {conj_col}")
            speaker.say(cn_phrase)
            time.sleep(1.2)  # 给用户反应时间：听到中文后回忆西语
            speaker.say(f"{subj_es} {conj}")
            time.sleep(0.2)
    print(Fore.GREEN + "  [OK] 乱序朗读完毕！")

# ============================================================
#  浏览模式
# ============================================================

def browse_mode(speaker: Speaker):
    # 星标动词优先排列
    browse_list = []
    for v in VERBS:
        if v['infinitive'] in starred_verbs:
            browse_list.append(v)
    for v in VERBS:
        if v['infinitive'] not in starred_verbs:
            browse_list.append(v)
    total = len(browse_list)
    idx = 0
    while True:
        clear_screen()
        print_title("1. 浏览模式 -- 逐个动词学习")
        star_info = ""
        if browse_list[idx]['infinitive'] in starred_verbs:
            star_info = f"  {Fore.YELLOW + Style.BRIGHT}⭐ 星标复习重点{Style.RESET_ALL}"
        print(f"  当前: 第 {idx+1}/{total} 个动词{star_info}")
        if starred_verbs:
            print(f"  (星标动词已置顶，共 {Fore.YELLOW}{len([v for v in browse_list if v['infinitive'] in starred_verbs])}{Style.RESET_ALL} 个)")
        print(f"  快捷键: [N]下一个 [P]上一个 [R]重读 [V]读变位表 [S]乱序读 [*]切换星标 [Q]返回")
        print_divider()

        verb = browse_list[idx]
        print_verb_header(verb, idx, total)
        print_conjugation_table(verb)

        # 朗读动词原形 + 中文
        print(Fore.GREEN + ">> 朗读动词原形和意思...")
        speaker.say(verb['infinitive'])
        time.sleep(0.3)
        speaker.say(verb['cn'])

        while True:
            cmd = input(Fore.WHITE + "\n  >> 请输入命令: ").strip().lower()
            if cmd in ('n', 'next', ''):
                idx = (idx + 1) % total
                break
            elif cmd in ('p', 'prev'):
                idx = (idx - 1) % total
                break
            elif cmd in ('r', 'repeat'):
                print(Fore.GREEN + ">> 重新朗读...")
                speaker.say(verb['infinitive'])
                time.sleep(0.3)
                speaker.say(verb['cn'])
            elif cmd in ('v', 'voice'):
                speak_conjugations(speaker, verb)
            elif cmd in ('s', 'shuffle'):
                print()
                print(Fore.WHITE + "  乱序朗读模式：")
                print("    Y — 西语 + 中文（如: él habla → 他说话）")
                print("    N — 仅西语连读（如: él habla，无停顿）")
                print("    P — 仅西语停顿（如: él … habla，主语后停顿）")
                print("    C — 中文先出（如: 他说话 … él habla，给你反应时间）")
                ans = input(Fore.WHITE + "  >> 请选择 (Y/N/P/C): ").strip().lower()
                if ans in ('y', 'yes', ''):
                    speak_conjugations_shuffle(speaker, verb, 'cn')
                elif ans in ('p', 'pause'):
                    speak_conjugations_shuffle(speaker, verb, 'pause')
                elif ans in ('c', 'cn', 'cnfirst'):
                    speak_conjugations_shuffle(speaker, verb, 'cnfirst')
                else:
                    speak_conjugations_shuffle(speaker, verb, 'fast')
            elif cmd in ('*', 'star', 'u', 'unstar'):
                inf = verb['infinitive']
                if inf in starred_verbs:
                    starred_verbs.discard(inf)
                    save_starred()
                    print(Fore.YELLOW + f"     ⭐ 已取消星标: {inf}")
                else:
                    starred_verbs.add(inf)
                    save_starred()
                    print(Fore.YELLOW + Style.BRIGHT + f"     ⭐ 已添加星标: {inf}")
                # 刷新浏览列表（星标动词置顶）
                browse_list = []
                for v in VERBS:
                    if v['infinitive'] in starred_verbs:
                        browse_list.append(v)
                for v in VERBS:
                    if v['infinitive'] not in starred_verbs:
                        browse_list.append(v)
                total = len(browse_list)
                # 找到当前动词在新列表中的位置
                for new_i, v in enumerate(browse_list):
                    if v['infinitive'] == inf:
                        idx = new_i
                        break
                break
            elif cmd == 'q':
                return
            else:
                print(Fore.RED + "  [?] 未知命令，请使用 N/P/R/V/S/*/Q")
# ============================================================
#  测验模式
# ============================================================

def normalize_accent(s: str) -> str:
    """去掉西班牙语重音符号，用于容错匹配（á→a, é→e, í→i, ó→o, ú→u, ü→u, ñ→n）"""
    replacements = str.maketrans('áéíóúüñÁÉÍÓÚÜÑ', 'aeiouunAEIOUUN')
    return s.translate(replacements).lower()

def quiz_mode(speaker: Speaker):
    clear_screen()
    print_title("2. 测验模式 -- 听读音回答")

    print("\n  规则说明：")
    print("  - 你会听到一个主语 + 动词变位（如 'él habla'、'nosotras comemos'）")
    print("  - 请输入: 动词原形 主语人称")
    print("  - 主语人称可以写西班牙语（él, ella, yo...）或中文（他, 她, 我...）")
    print("  - 例如听到 'ella habla'，回答: hablar ella  或  hablar 她")
    print("  - 也可以输入 R 重听  Q 退出测验")
    print()

    print("  主语人称参考:")
    print("  ┌─────────────────────────┬──────────────────────────┐")
    print("  │ 西班牙语                 │ 中文                      │")
    print("  ├─────────────────────────┼──────────────────────────┤")
    print("  │ yo                      │ 我                       │")
    print("  │ tú                      │ 你                       │")
    print("  │ él / ella / usted       │ 他 / 她 / 您              │")
    print("  │ nosotros / nosotras     │ 我们                     │")
    print("  │ vosotros / vosotras     │ 你们                     │")
    print("  │ ellos / ellas / ustedes │ 他们 / 她们 / 诸位        │")
    print("  └─────────────────────────┴──────────────────────────┘")
    print()

    print("  请选择题库范围：")
    print("  1. 所有动词 (14个)")
    print("  2. 仅规则动词 (7个)")
    print("  3. 仅不规则动词 (7个)")
    print("  4. 自定义数量")
    scope = input(Fore.WHITE + "\n  >> 请选择 (1-4): ").strip()

    if scope == '2':
        pool = [v for v in VERBS if '规则' in v['category']]
    elif scope == '3':
        pool = [v for v in VERBS if '不规则' in v['category']]
    elif scope == '4':
        try:
            n = int(input("  出题数量: ").strip())
            pool = random.choices(VERBS, k=n)
        except:
            pool = VERBS
    else:
        pool = VERBS

    print()
    cn_choice = input(Fore.WHITE + "  >> 是否包含中文翻译？(Y/N，默认 N): ").strip().lower()
    include_cn = cn_choice in ('y', 'yes')

    random.shuffle(pool)
    questions = []
    for v in pool:
        i = random.randint(0, 5)
        questions.append((v, i))

    score = 0
    total = len(questions)

    for qi, (verb, conj_idx) in enumerate(questions):
        clear_screen()
        print_title(f"2. 测验模式 -- 第 {qi+1}/{total} 题  (得分: {score})")
        print_divider()
        print(Fore.WHITE + "  主语参考: yo=我 tú=你 él=他 ella=她 usted=您 nosotros=我们 nosotras=我们 vosotros=你们 vosotras=你们 ellos=他们 ellas=她们 ustedes=诸位")
        print_divider()

        subj_display, subj_cn = SUBJECTS[conj_idx]
        conj = verb['conjugations'][conj_idx]
        # 从复合主语中随机选一个来读（如 él/ella/usted 随机抽一个）
        spoken_subject, spoken_cn = pick_random_subject(conj_idx)
        phrase = f"{spoken_subject} {conj}"

        # 朗读主语+变位 连读
        print(Fore.GREEN + f">> 请听题...")
        speaker.say(phrase)

        prompt_text = "动词原形 主语人称 中文意思" if include_cn else "动词原形 主语人称"
        answer = input(Fore.WHITE + f"\n  >> 请输入 [{prompt_text}] (R重听 Q退出): ").strip()

        if answer.lower() == 'q':
            break
        elif answer.lower() == 'r':
            speaker.say(phrase)
            answer = input(Fore.WHITE + f"  >> 请输入 [{prompt_text}]: ").strip()

        if answer.lower() == 'q':
            break

        parts = answer.strip().split()
        correct_inf = verb['infinitive']

        user_inf = parts[0].lower() if len(parts) >= 1 else ''
        user_pronoun_raw = parts[1] if len(parts) >= 2 else ''
        user_cn_raw = parts[2] if len(parts) >= 3 else ''

        inf_ok = (normalize_accent(user_inf) == normalize_accent(correct_inf))
        pronoun_ok = (normalize_accent(user_pronoun_raw) == normalize_accent(spoken_subject)) or (user_pronoun_raw == spoken_cn)
        correct = inf_ok and pronoun_ok

        cn_phrase_q = f"{spoken_cn}{verb['cn']}"
        if correct:
            print(Fore.GREEN + Style.BRIGHT + f"\n  [OK] 正确！{verb['infinitive']} ({verb['cn']}) -- {spoken_subject} {conj} → {cn_phrase_q}")
            score += 1
        else:
            print(Fore.RED + Style.BRIGHT + f"\n  [X] 错误！")
            if not inf_ok:
                print(Fore.RED + f"     动词原形: 你答 '{user_inf}'，正确是 '{correct_inf}'")
            if not pronoun_ok:
                print(Fore.RED + f"     主语人称: 你答 '{user_pronoun_raw}'，正确是 '{spoken_subject}' 或 '{spoken_cn}'")
            print(Fore.YELLOW + f"     正确答案: {verb['infinitive']} ({verb['cn']}) -- {spoken_subject} {conj} → {cn_phrase_q}")

        # 中文翻译：自检模式，不计分
        if include_cn:
            if user_cn_raw:
                print(Fore.CYAN + f"     你的翻译: {user_cn_raw}  |  参考翻译: {verb['cn']}")
            else:
                print(Fore.CYAN + f"     参考翻译: {verb['cn']}")

        # 写错了询问是否星标
        if not correct:
            inf = verb['infinitive']
            if inf in starred_verbs:
                print(Fore.YELLOW + f"     ⭐ 已星标复习重点")
            else:
                star_ans = input(Fore.WHITE + "  >> 星标该动词作为复习重点？(Y/N，默认 N): ").strip().lower()
                if star_ans in ('y', 'yes'):
                    starred_verbs.add(inf)
                    save_starred()
                    print(Fore.YELLOW + f"     ⭐ 已添加星标: {inf}")
        else:
            # 答对了，显示星标状态
            inf = verb['infinitive']
            if inf in starred_verbs:
                print(Fore.GREEN + f"     ⭐ 已星标（你之前标记的复习重点）")

        print()
        print_conjugation_table(verb, highlight_idx=conj_idx)

        if qi < total - 1:
            input(Fore.WHITE + "\n  按 Enter 继续下一题: ")

    clear_screen()
    print_title("2. 测验结果")
    pct = score / total * 100 if total > 0 else 0
    print(f"\n  得分: {Fore.GREEN + str(score)}/{total}  ({pct:.0f}%)")
    if pct == 100:
        print(Fore.GREEN + Style.BRIGHT + "  满分！太棒了！¡Excelente!")
    elif pct >= 80:
        print(Fore.CYAN + "  很不错！¡Muy bien!")
    elif pct >= 60:
        print(Fore.YELLOW + "  还需要多练习哦！")
    else:
        print(Fore.RED + "  继续加油！¡Ánimo!")
    input(Fore.WHITE + "\n  按 Enter 返回主菜单: ")

# ============================================================
#  听写变位模式
# ============================================================

def dictation_mode(speaker: Speaker):
    clear_screen()
    print_title("3. 听写模式 -- 听原形+人称，写出变位")

    print("\n  规则说明：")
    print("  - 你会听到一个动词原形，然后是一个人称呼（如 'tener ... yo'）")
    print("  - 请写出该动词在该人称下的变位形式")
    print("  - 例如听到 'tener ... yo'，回答: tengo")
    print("  - 也可以输入 R 重听  Q 退出")
    print()

    print("  请选择题库范围：")
    print("  1. 所有动词 (14个)")
    print("  2. 仅规则动词 (7个)")
    print("  3. 仅不规则动词 (7个)")
    print("  4. 自定义数量")
    scope = input(Fore.WHITE + "\n  >> 请选择 (1-4): ").strip()

    if scope == '2':
        pool = [v for v in VERBS if '规则' in v['category']]
    elif scope == '3':
        pool = [v for v in VERBS if '不规则' in v['category']]
    elif scope == '4':
        try:
            n = int(input("  出题数量: ").strip())
            pool = random.choices(VERBS, k=n)
        except:
            pool = VERBS
    else:
        pool = VERBS

    questions = []
    for v in pool:
        i = random.randint(0, 5)
        questions.append((v, i))
    random.shuffle(questions)

    score = 0
    total = len(questions)

    for qi, (verb, conj_idx) in enumerate(questions):
        clear_screen()
        print_title(f"3. 听写模式 -- 第 {qi+1}/{total} 题  (得分: {score})")
        print_divider()

        conj = verb['conjugations'][conj_idx]
        spoken_subject, spoken_cn = pick_random_subject(conj_idx)

        print(Fore.YELLOW + f"  动词原形: {verb['infinitive']} ({verb['cn']})")
        print(Fore.CYAN + f"  人称: {spoken_subject} ({spoken_cn})")
        print(Fore.GREEN + f">> 请听题...")
        speaker.say(verb['infinitive'])
        time.sleep(0.5)
        speaker.say(spoken_subject)

        answer = input(Fore.WHITE + "\n  >> 请输入变位形式 (R重听 Q退出): ").strip()

        if answer.lower() == 'q':
            break
        elif answer.lower() == 'r':
            speaker.say(verb['infinitive'])
            time.sleep(0.5)
            speaker.say(spoken_subject)
            answer = input(Fore.WHITE + "  >> 请输入变位形式: ").strip()

        if answer.lower() == 'q':
            break

        user_conj = answer.strip()
        correct_conj = conj

        conj_ok = (normalize_accent(user_conj) == normalize_accent(correct_conj))

        if conj_ok:
            print(Fore.GREEN + Style.BRIGHT + f"\n  [OK] 正确！{spoken_subject} {correct_conj}")
            score += 1
        else:
            print(Fore.RED + Style.BRIGHT + f"\n  [X] 错误！")
            print(Fore.RED + f"     你写: '{user_conj}'")
            print(Fore.GREEN + f"     正确: '{correct_conj}'")
            # 分析差异
            diff_note = ""
            if normalize_accent(user_conj[:-1]) == normalize_accent(correct_conj[:-1]):
                diff_note = "（词干对了，注意词尾！）"
            elif len(user_conj) == len(correct_conj):
                diff_note = "（注意元音变化！）"
            if diff_note:
                print(Fore.YELLOW + f"     💡 {diff_note}")

        print_conjugation_table(verb, highlight_idx=conj_idx)

        if qi < total - 1:
            input(Fore.WHITE + "\n  按 Enter 继续下一题: ")

    clear_screen()
    print_title("3. 听写测验结果")
    pct = score / total * 100 if total > 0 else 0
    print(f"\n  得分: {Fore.GREEN + str(score)}/{total}  ({pct:.0f}%)")
    if pct == 100:
        print(Fore.GREEN + Style.BRIGHT + "  满分！拼写完美！¡Perfecto!")
    elif pct >= 80:
        print(Fore.CYAN + "  很不错！多注意元音变化！")
    elif pct >= 60:
        print(Fore.YELLOW + "  还需要多写多练哦！¡Practica más!")
    else:
        print(Fore.RED + "  继续加油！拼写需要反复练习！¡Ánimo!")
    input(Fore.WHITE + "\n  按 Enter 返回主菜单: ")

# ============================================================
#  设置
# ============================================================

def settings_menu(speaker: Speaker):
    while True:
        clear_screen()
        print_title("4. 设置")
        if speaker.tts_enabled:
            print(f"  当前语速: {Fore.GREEN}{speaker.get_rate()}  (-10 ~ 10)")
            print(f"  当前音量: {Fore.GREEN}{speaker.get_volume()}  (0 ~ 100)")
        else:
            print(Fore.RED + "  TTS 引擎未启用")
        print()
        print("  1. 调整语速 (-10 ~ 10, 默认 0)")
        print("  2. 调整音量 (0 ~ 100, 默认 100)")
        print("  3. 列出所有语音")
        print("  4. 切换语音")
        print("  Q. 返回主菜单")

        cmd = input(Fore.WHITE + "\n  >> 请选择: ").strip().lower()
        if cmd == '1':
            try:
                rate = int(input("  输入语速 (-10 ~ 10): ").strip())
                rate = max(-10, min(10, rate))
                speaker.set_rate(rate)
                print(Fore.GREEN + f"  [OK] 语速已设为 {rate}")
            except:
                print(Fore.RED + "  输入无效")
            input("  按 Enter 继续: ")
        elif cmd == '2':
            try:
                vol = int(input("  输入音量 (0 ~ 100): ").strip())
                vol = max(0, min(100, vol))
                speaker.set_volume(vol)
                print(Fore.GREEN + f"  [OK] 音量已设为 {vol}")
            except:
                print(Fore.RED + "  输入无效")
            input("  按 Enter 继续: ")
        elif cmd == '3':
            print("\n  可用语音列表：")
            speaker.list_voices()
            input("\n  按 Enter 继续: ")
        elif cmd == '4':
            speaker.list_voices()
            try:
                idx = int(input("  输入语音序号: ").strip())
                speaker.set_voice(idx)
                print(Fore.GREEN + "  [OK] 语音已切换")
            except:
                print(Fore.RED + "  输入无效")
            input("  按 Enter 继续: ")
        elif cmd == 'q':
            return

# ============================================================
#  主菜单
# ============================================================

def main_menu(speaker: Speaker):
    while True:
        clear_screen()
        print_title("西班牙语动词变位 -- 听觉主导学习工具")
        print(f"  共 {Fore.GREEN + str(len(VERBS))} 个动词")
        print(f"    · 规则 -ar 动词: {Fore.CYAN}3 个  (hablar, cantar, estudiar)")
        print(f"    · 规则 -er 动词: {Fore.CYAN}2 个  (comer, beber)")
        print(f"    · 规则 -ir 动词: {Fore.CYAN}2 个  (cumplir, subir)")
        print(f"    · 不规则动词:   {Fore.CYAN}7 个  (ser, estar, tener, poder, poner, venir, ir)")
        if not speaker.tts_enabled:
            print()
            print(Fore.RED + Style.BRIGHT + "  [!] 语音引擎未就绪！pip install pywin32")
        print()
        print_divider()
        print()
        print("  " + Fore.GREEN + Style.BRIGHT + "1. 浏览模式" + Fore.WHITE + "  -- 逐个浏览动词、听发音（不规则高亮）")
        print("  " + Fore.GREEN + Style.BRIGHT + "2. 测验模式" + Fore.WHITE + "  -- 听读音，回答动词原形+主语人称(可选中文)")
        print("  " + Fore.GREEN + Style.BRIGHT + "3. 听写模式" + Fore.WHITE + "  -- 听原形+人称，写出变位形式")
        print("  " + Fore.GREEN + Style.BRIGHT + "4. 设置" + Fore.WHITE + "      -- 语速/音量/语音")
        print("  " + Fore.GREEN + Style.BRIGHT + "Q. 退出")
        print()

        cmd = input(Fore.WHITE + "  >> 请选择模式: ").strip().lower()

        if cmd == '1':
            browse_mode(speaker)
        elif cmd == '2':
            quiz_mode(speaker)
        elif cmd == '3':
            dictation_mode(speaker)
        elif cmd == '4':
            settings_menu(speaker)
        elif cmd == 'q':
            clear_screen()
            print(Fore.YELLOW + Style.BRIGHT + "\n  ¡Adiós! 再见！下次继续加油！\n")
            sys.exit(0)
        else:
            print(Fore.RED + "  [?] 无效选择，请重试")
            time.sleep(0.8)

# ============================================================
#  入口
# ============================================================

if __name__ == '__main__':
    try:
        load_starred()
        print("正在初始化 SAPI 双语音引擎...")
        speaker = Speaker()
        if speaker.tts_enabled:
            es_name = speaker.voice_es.GetDescription() if speaker.voice_es else "N/A"
            cn_name = speaker.voice_cn.GetDescription() if speaker.voice_cn else "N/A"
            print(Fore.GREEN + f"  [OK] 西语: {es_name}")
            print(Fore.GREEN + f"  [OK] 中文: {cn_name}")
            print(Fore.CYAN + "  (朗读时自动切换，中文用中文语音，西语用西语语音)")
        main_menu(speaker)
    except KeyboardInterrupt:
        print(Fore.YELLOW + "\n\n  ¡Hasta luego! 再见！\n")
        sys.exit(0)
    except Exception as e:
        err_msg = traceback.format_exc()
        log_error(err_msg)
        print(Fore.RED + f"\n  [X] 发生错误:")
        print(Fore.RED + f"      {e}")
        print(Fore.YELLOW + f"\n  详细错误已写入: {LOG_FILE}")
        input("\n  按 Enter 退出: ")
        sys.exit(1)
