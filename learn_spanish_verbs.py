#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
西班牙语动词变位 —— 听觉主导学习工具
=============================================
功能：
  1. 浏览模式 — 逐个动词浏览，听+看 6 种人称变位
  2. 学习模式 — 按类别系统学习，自动朗读每个变位
  3. 测验模式 — 听读音，回答动词原形+人称，计分
  4. 拼写模式 — 听动词原形，默写拼写和中文意思

依赖：pywin32, colorama
安装：pip install pywin32 colorama
"""

import sys
import os
import time
import random
import traceback

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_error_log.txt')

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
    },
    {
        "infinitive": "estar",
        "cn": "处于、处在",
        "category": "不规则动词",
        "conjugations": ["estoy", "estás", "está", "estamos", "estáis", "están"],
        "example": "",
    },
    {
        "infinitive": "tener",
        "cn": "有",
        "category": "不规则动词",
        "conjugations": ["tengo", "tienes", "tiene", "tenemos", "tenéis", "tienen"],
        "example": "",
    },
    {
        "infinitive": "poder",
        "cn": "能够",
        "category": "不规则动词",
        "conjugations": ["puedo", "puedes", "puede", "podemos", "podéis", "pueden"],
        "example": "",
    },
    {
        "infinitive": "poner",
        "cn": "放",
        "category": "不规则动词",
        "conjugations": ["pongo", "pones", "pone", "ponemos", "ponéis", "ponen"],
        "example": "",
    },
    {
        "infinitive": "venir",
        "cn": "来",
        "category": "不规则动词",
        "conjugations": ["vengo", "vienes", "viene", "venimos", "venís", "vienen"],
        "example": "",
    },
    {
        "infinitive": "ir",
        "cn": "去",
        "category": "不规则动词",
        "conjugations": ["voy", "vas", "va", "vamos", "vais", "van"],
        "example": "",
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
    print(Fore.MAGENTA + Style.BRIGHT + f"--- [{idx+1}/{total}] ---")
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

def print_conjugation_table(verb: dict):
    print()
    print(Fore.WHITE + Style.BRIGHT + f"  {'主格人称代词':<26} {'变位':<16} {'西 → 中':<20}")
    print(Fore.WHITE + "  " + "-" * 64)
    for i, (subj_es, subj_cn) in enumerate(SUBJECTS):
        conj = verb['conjugations'][i]
        cn_phrase = make_cn_phrase(i, 0, verb['cn'])
        print(f"  {Fore.CYAN + subj_es:<26} {Fore.YELLOW + Style.BRIGHT + conj:<16} {Fore.GREEN + conj:<12} → {Fore.WHITE + cn_phrase}")
    print()

def print_divider():
    print(Fore.BLUE + "-" * 60)

def speak_conjugations(speaker: Speaker, verb: dict):
    """朗读全部变位 —— 西语 + 中文都念，每个子主语配匹配的中文"""
    print(Fore.GREEN + ">> 朗读全部变位（西语 → 中文，复合主语已拆开）...")
    for i, (subj_display, _) in enumerate(SUBJECT_SPLIT):
        conj = verb['conjugations'][i]
        parts_es = SUBJECT_SPLIT[i][1]
        parts_cn = SUBJECTS_CN_SPLIT[i]
        for j, subj_es in enumerate(parts_es):
            cn_phrase = make_cn_phrase(i, j, verb['cn'])
            es_phrase = f"{subj_es} {conj}"
            print(f"     {Fore.CYAN + subj_es:<14} {Fore.YELLOW + conj:<13} → {Fore.GREEN + cn_phrase}")
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
    all_pairs = []
    for i in range(6):
        conj = verb['conjugations'][i]
        parts_es = SUBJECT_SPLIT[i][1]
        parts_cn = SUBJECTS_CN_SPLIT[i]
        for j, subj_es in enumerate(parts_es):
            cn_subj = parts_cn[j]
            cn_phrase = f"{cn_subj}{verb['cn']}"
            all_pairs.append((subj_es, conj, cn_phrase))
    random.shuffle(all_pairs)

    mode_desc = {
        "cn": "西语 + 中文",
        "fast": "仅西语（连读）",
        "pause": "仅西语（主语-停顿-变位）",
        "cnfirst": "中文先出（等反应）→ 西语验证"
    }[mode]
    print(Fore.GREEN + f">> 乱序朗读全部变位（{mode_desc}）...")
    for subj_es, conj, cn_phrase in all_pairs:
        if mode == 'cn':
            print(f"     {Fore.CYAN + subj_es:<14} {Fore.YELLOW + conj:<13} → {Fore.GREEN + cn_phrase}")
            speaker.say(f"{subj_es} {conj}")
            time.sleep(0.15)
            speaker.say(cn_phrase)
            time.sleep(0.2)
        elif mode == 'fast':
            print(f"     {Fore.CYAN + subj_es:<14} {Fore.YELLOW + conj}")
            speaker.say(f"{subj_es} {conj}")
            time.sleep(0.12)
        elif mode == 'pause':
            print(f"     {Fore.CYAN + subj_es:<14} {Fore.YELLOW + conj}")
            speaker.say(subj_es)
            time.sleep(0.35)
            speaker.say(conj)
            time.sleep(0.15)
        elif mode == 'cnfirst':
            print(f"     {Fore.GREEN + cn_phrase:<12} → {Fore.CYAN + subj_es:<14} {Fore.YELLOW + conj}")
            speaker.say(cn_phrase)
            time.sleep(1.2)  # 给用户反应时间：听到中文后回忆西语
            speaker.say(f"{subj_es} {conj}")
            time.sleep(0.2)
    print(Fore.GREEN + "  [OK] 乱序朗读完毕！")

# ============================================================
#  浏览模式
# ============================================================

def browse_mode(speaker: Speaker):
    total = len(VERBS)
    idx = 0
    while True:
        clear_screen()
        print_title("1. 浏览模式 -- 逐个动词学习")
        print(f"  当前: 第 {idx+1}/{total} 个动词")
        print(f"  快捷键: [N]下一个 [P]上一个 [R]重读 [V]读变位表 [S]乱序读 [Q]返回")
        print_divider()

        verb = VERBS[idx]
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
            elif cmd == 'q':
                return
            else:
                print(Fore.RED + "  [?] 未知命令，请使用 N/P/R/V/S/Q")

# ============================================================
#  学习模式
# ============================================================

def study_mode(speaker: Speaker):
    categories = {}
    for v in VERBS:
        cat = v['category']
        categories.setdefault(cat, []).append(v)

    cat_names = list(categories.keys())
    cat_idx = 0

    while True:
        clear_screen()
        print_title("2. 学习模式 -- 按类别系统学习")

        cat = cat_names[cat_idx]
        verbs_in_cat = categories[cat]
        print(f"\n  当前类别: {Fore.GREEN + Style.BRIGHT + cat}")
        print(f"     包含 {len(verbs_in_cat)} 个动词")
        print(f"\n  [N]下一类  [P]上一类  [S]开始学习本类  [Q]返回")

        cmd = input(Fore.WHITE + "\n  >> 请输入命令: ").strip().lower()
        if cmd in ('n', ''):
            cat_idx = (cat_idx + 1) % len(cat_names)
            continue
        elif cmd == 'p':
            cat_idx = (cat_idx - 1) % len(cat_names)
            continue
        elif cmd == 'q':
            return
        elif cmd == 's':
            _study_category(speaker, cat, verbs_in_cat)
        else:
            print(Fore.RED + "  [?] 未知命令")

def _study_category(speaker: Speaker, cat_name: str, verbs_in_cat: list):
    for vi, verb in enumerate(verbs_in_cat):
        clear_screen()
        print_title(f"2. 学习模式 -- {cat_name}")
        print(f"  动词 {vi+1}/{len(verbs_in_cat)}")
        print_divider()
        print_verb_header(verb, vi, len(verbs_in_cat))
        print_conjugation_table(verb)

        # 朗读原形 + 中文
        print(Fore.GREEN + ">> 朗读动词原形和意思...")
        speaker.say(verb['infinitive'])
        time.sleep(0.3)
        speaker.say(verb['cn'])
        time.sleep(0.3)

        # 朗读 6 个变位 —— 主语+变位 连读
        speak_conjugations(speaker, verb)

        print(Fore.GREEN + "\n  [OK] 本动词学习完毕！")
        if verb.get('example'):
            print(Fore.YELLOW + f"  例句: {verb['example']}")
            speaker.say(verb['example'])

        if vi < len(verbs_in_cat) - 1:
            cmd = input(Fore.WHITE + "\n  按 Enter 继续下一个动词 (Q 退出): ").strip().lower()
            if cmd == 'q':
                return
        else:
            input(Fore.WHITE + "\n  本类别学习完毕！按 Enter 返回: ")

# ============================================================
#  测验模式
# ============================================================

def quiz_mode(speaker: Speaker):
    clear_screen()
    print_title("3. 测验模式 -- 听读音回答")

    print("\n  规则说明：")
    print("  - 你会听到一个主语 + 动词变位（如 'él habla'、'nosotras comemos'）")
    print("  - 请输入: 动词原形 人称(序号 1-6)")
    print("  - 人称序号: 1=yo  2=tú  3=él/ella/usted  4=nosotros/as  5=vosotros/as  6=ellos/ellas/ustedes")
    print("  - 例如听到 'ella habla'，回答: hablar 3")
    print("  - 也可以直接输入 q 退出测验")
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

    random.shuffle(pool)
    questions = []
    for v in pool:
        i = random.randint(0, 5)
        questions.append((v, i))

    score = 0
    total = len(questions)

    for qi, (verb, conj_idx) in enumerate(questions):
        clear_screen()
        print_title(f"3. 测验模式 -- 第 {qi+1}/{total} 题  (得分: {score})")
        print_divider()

        subj_display, subj_cn = SUBJECTS[conj_idx]
        conj = verb['conjugations'][conj_idx]
        # 从复合主语中随机选一个来读（如 él/ella/usted 随机抽一个）
        spoken_subject, spoken_cn = pick_random_subject(conj_idx)
        phrase = f"{spoken_subject} {conj}"

        # 朗读主语+变位 连读
        print(Fore.GREEN + f">> 请听题...")
        speaker.say(phrase)

        answer = input(Fore.WHITE + "\n  >> 请输入 [动词原形 人称1-6] (R重听 Q退出): ").strip()

        if answer.lower() == 'q':
            break
        elif answer.lower() == 'r':
            speaker.say(phrase)
            answer = input(Fore.WHITE + "  >> 请输入 [动词原形 人称1-6]: ").strip()

        if answer.lower() == 'q':
            break

        parts = answer.strip().split()
        correct_inf = verb['infinitive']
        correct_person = str(conj_idx + 1)

        user_inf = parts[0].lower() if len(parts) >= 1 else ''
        user_person = parts[1] if len(parts) >= 2 else ''

        inf_ok = (user_inf == correct_inf)
        person_ok = (user_person == correct_person)
        correct = inf_ok and person_ok

        cn_phrase_q = f"{spoken_cn}{verb['cn']}"
        if correct:
            print(Fore.GREEN + Style.BRIGHT + f"\n  [OK] 正确！{verb['infinitive']} ({verb['cn']}) -- {spoken_subject} {conj} → {cn_phrase_q}")
            score += 1
        else:
            print(Fore.RED + Style.BRIGHT + f"\n  [X] 错误！")
            if not inf_ok:
                print(Fore.RED + f"     动词原形: 你答 '{user_inf}'，正确是 '{correct_inf}'")
            if not person_ok:
                print(Fore.RED + f"     人称: 你答 '{user_person}'，正确是 '{correct_person}' ({subj_display})")
            print(Fore.YELLOW + f"     正确答案: {verb['infinitive']} ({verb['cn']}) -- {spoken_subject} {conj} → {cn_phrase_q}")

        print()
        print(Fore.WHITE + Style.BRIGHT + f"  {'主格人称代词':<26} {'变位':<16} {'西 → 中'}")
        print(Fore.WHITE + "  " + "-" * 64)
        for j, (s_es, s_cn) in enumerate(SUBJECTS):
            marker = " <-- 本题" if j == conj_idx else ""
            cn_p = make_cn_phrase(j, 0, verb['cn'])
            print(f"  {Fore.CYAN + s_es:<26} {Fore.YELLOW + verb['conjugations'][j]:<16} {Fore.GREEN + verb['conjugations'][j]:<12} → {Fore.WHITE + cn_p}" + marker)

        if qi < total - 1:
            input(Fore.WHITE + "\n  按 Enter 继续下一题: ")

    clear_screen()
    print_title("3. 测验结果")
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
#  拼写模式
# ============================================================

def spelling_mode(speaker: Speaker):
    clear_screen()
    print_title("4. 拼写模式 -- 听发音默写")

    print("\n  规则说明：")
    print("  - 你会听到一个动词原形的发音")
    print("  - 请输入: 拼写 中文意思")
    print("  - 例如听到 'hablar'，回答: hablar 说话")
    print()

    pool = VERBS.copy()
    random.shuffle(pool)

    score = 0
    total = len(pool)

    for qi, verb in enumerate(pool):
        clear_screen()
        print_title(f"4. 拼写模式 -- 第 {qi+1}/{total} 题  (得分: {score})")
        print_divider()

        print(Fore.GREEN + ">> 请听动词发音...")
        speaker.say(verb['infinitive'])

        print(Fore.CYAN + f"  提示: 属于「{verb['category']}」")
        print()

        answer = input(Fore.WHITE + "  >> 请输入 [拼写 中文意思] (R重听 Q退出): ").strip()

        if answer.lower() == 'q':
            break
        elif answer.lower() == 'r':
            speaker.say(verb['infinitive'])
            answer = input(Fore.WHITE + "  >> 请输入 [拼写 中文意思]: ").strip()

        if answer.lower() == 'q':
            break

        parts = answer.strip().split(maxsplit=1)
        user_spelling = parts[0].lower() if len(parts) >= 1 else ''
        user_cn = parts[1] if len(parts) >= 2 else ''

        spell_ok = (user_spelling == verb['infinitive'])
        cn_ok = (user_cn == verb['cn'])

        if spell_ok and cn_ok:
            print(Fore.GREEN + Style.BRIGHT + f"\n  [OK] 完全正确！{verb['infinitive']} -- {verb['cn']}")
            score += 1
        else:
            if not spell_ok:
                print(Fore.RED + f"  [X] 拼写错误: 你写 '{user_spelling}'，正确是 '{verb['infinitive']}'")
            if not cn_ok:
                print(Fore.RED + f"  [X] 意思错误: 你写 '{user_cn}'，正确是 '{verb['cn']}'")
            print(Fore.YELLOW + f"     正确答案: {verb['infinitive']} -- {verb['cn']}  ({verb['category']})")

        if spell_ok and not cn_ok:
            score += 0.5

        if qi < total - 1:
            input(Fore.WHITE + "\n  按 Enter 继续下一题: ")

    clear_screen()
    print_title("4. 拼写测验结果")
    print(f"\n  得分: {Fore.GREEN + str(score)}/{total}")
    input(Fore.WHITE + "\n  按 Enter 返回主菜单: ")

# ============================================================
#  设置
# ============================================================

def settings_menu(speaker: Speaker):
    while True:
        clear_screen()
        print_title("5. 设置")
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
        print("  " + Fore.GREEN + Style.BRIGHT + "1. 浏览模式" + Fore.WHITE + "  -- 逐个浏览动词、听发音")
        print("  " + Fore.GREEN + Style.BRIGHT + "2. 学习模式" + Fore.WHITE + "  -- 按类别自动系统学习")
        print("  " + Fore.GREEN + Style.BRIGHT + "3. 测验模式" + Fore.WHITE + "  -- 听读音，回答动词原形+人称")
        print("  " + Fore.GREEN + Style.BRIGHT + "4. 拼写模式" + Fore.WHITE + "  -- 听发音，默写拼写和意思")
        print("  " + Fore.GREEN + Style.BRIGHT + "5. 设置" + Fore.WHITE + "      -- 语速/音量/语音")
        print("  " + Fore.GREEN + Style.BRIGHT + "Q. 退出")
        print()

        cmd = input(Fore.WHITE + "  >> 请选择模式: ").strip().lower()

        if cmd == '1':
            browse_mode(speaker)
        elif cmd == '2':
            study_mode(speaker)
        elif cmd == '3':
            quiz_mode(speaker)
        elif cmd == '4':
            spelling_mode(speaker)
        elif cmd == '5':
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
