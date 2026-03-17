from enum import IntEnum
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Callable, Tuple
import csv
import os
import threading


# ==========================================
#  定数
# ==========================================

EFFECT_TYPES = [
    "未設定",
    "敏捷強化", "筋力強化", "知力強化",
    "特攻ダメージ強化", "特攻回復強化", "マスタリー回復強化", "精鋭打撃",
    "魔法耐性", "物理耐性",
    "集中・攻撃速度", "集中・詠唱", "集中・会心", "集中・幸運",
    "極・ダメージ増強", "極・適応力", "極・HP凝縮", "極・応急処置",
    "極・絶境守護", "極・HP変動", "極・HP吸収", "極・幸運会心",
]

# "未設定"を除いた効果タイプリスト（サマリー用）
MODULE_EFFECT_NAMES = EFFECT_TYPES[1:]

THRESHOLDS = [1, 4, 8, 12, 16, 20]


# ==========================================
#  Enum
# ==========================================

class SearchPriority(IntEnum):
    """検索条件の優先度"""
    Normal = 0     # 通常
    Priority = 1   # 優先 (Lv5/Lv6でボーナス)

    def __str__(self):
        labels = {0: "通常", 1: "優先"}
        return labels.get(self.value, str(self.value))


# ==========================================
#  データクラス
# ==========================================

@dataclass
class ModuleOption:
    """モジュールの1つの効果オプション（タイプ + 値）"""
    selected_type: str = "未設定"
    value: int = 0
    on_changed: Optional[Callable] = field(default=None, repr=False)

    def set_type(self, t: str):
        self.selected_type = t
        if self.on_changed:
            self.on_changed()

    def set_value(self, v: int):
        self.value = max(0, min(v, 10))
        if self.on_changed:
            self.on_changed()


@dataclass
class ModuleItem:
    """1つのモジュールカード"""
    name: str = "モジュール"
    is_equipped: bool = False
    options: List[ModuleOption] = field(default_factory=list)
    on_update: Optional[Callable] = field(default=None, repr=False)
    on_equip_changed: Optional[Callable] = field(default=None, repr=False)

    def __post_init__(self):
        if not self.options:
            self.options = [
                ModuleOption(on_changed=self.on_update),
                ModuleOption(on_changed=self.on_update),
                ModuleOption(on_changed=self.on_update),
            ]

    def set_equipped(self, val: bool):
        self.is_equipped = val
        if self.on_equip_changed:
            self.on_equip_changed(self)


@dataclass
class ModuleEffectStatus:
    """1つの効果タイプのステータス（サマリー表示用）"""
    name: str = ""
    current_total: int = 0
    level: int = 0
    next_threshold: int = 1

    @property
    def is_visible(self) -> bool:
        return self.level > 0 or self.current_total > 0


# ==========================================
#  モジュール効果計算
# ==========================================

@dataclass
class ModuleStats:
    """モジュール効果から得られるステータスボーナス"""
    phys_atk: int = 0       # 物理攻撃力
    mag_atk: int = 0        # 魔法攻撃力
    pm_atk: int = 0         # 物理/魔法攻撃力 (共通)
    strength: int = 0       # 筋力
    intellect: int = 0      # 知力
    agility: int = 0        # 敏捷
    sia: int = 0            # 筋力/知力/敏捷 (共通)
    max_hp: int = 0         # 最大HP
    endurance: int = 0      # 耐久力
    phys_def: int = 0       # 物理防御力
    all_attr: int = 0       # 全属性強度
    effects: List[str] = field(default_factory=list)  # ユニーク効果テキスト


def calculate_level(val: int) -> Tuple[int, int]:
    """ポイント値からレベルと次の閾値を返す"""
    lv = 0
    next_th = THRESHOLDS[0]
    for i, th in enumerate(THRESHOLDS):
        if val >= th:
            lv = i + 1
            next_th = THRESHOLDS[i + 1] if i + 1 < len(THRESHOLDS) else THRESHOLDS[-1]
        else:
            next_th = th
            break
    return lv, next_th


def update_module_effects(modules: List[ModuleItem], effect_summary: List[ModuleEffectStatus]) -> int:
    """装備中モジュールからエフェクトサマリーを更新し、リンク効果総計を返す"""
    type_totals: Dict[str, int] = {s.name: 0 for s in effect_summary}
    link_total = 0

    for mod in modules:
        if not mod.is_equipped:
            continue
        for opt in mod.options:
            if opt.selected_type != "未設定":
                if opt.selected_type in type_totals:
                    type_totals[opt.selected_type] += opt.value
                link_total += opt.value

    for status in effect_summary:
        val = type_totals.get(status.name, 0)
        status.current_total = val
        lv, next_th = calculate_level(val)
        status.level = lv
        status.next_threshold = next_th

    return link_total


def calculate_module_stats(effect_summary: List[ModuleEffectStatus], link_total: int) -> ModuleStats:
    """全エフェクトのレベルからステータスボーナスを計算する"""
    stats = ModuleStats()

    # リンク効果: 筋力/知力/敏捷 += linkTotal * 2
    stats.sia += link_total * 2

    for status in effect_summary:
        if status.level == 0:
            continue
        lv = status.level
        name = status.name

        if name == "敏捷強化":
            _apply_agility_boost(lv, stats)
        elif name == "筋力強化":
            _apply_strength_boost(lv, stats)
        elif name == "知力強化":
            _apply_intellect_boost(lv, stats)
        elif name == "特攻ダメージ強化":
            _apply_special_dmg_boost(lv, stats)
        elif name == "特攻回復強化":
            _apply_special_recovery_boost(lv, stats)
        elif name == "マスタリー回復強化":
            _apply_mastery_recovery_boost(lv, stats)
        elif name == "精鋭打撃":
            _apply_elite_strike(lv, stats)
        elif name == "魔法耐性":
            _apply_magic_resist(lv, stats)
        elif name == "物理耐性":
            _apply_phys_resist(lv, stats)
        elif name == "集中・攻撃速度":
            _apply_focus_atk_speed(lv, stats)
        elif name == "集中・詠唱":
            _apply_focus_cast(lv, stats)
        elif name == "集中・会心":
            _apply_focus_crit(lv, stats)
        elif name == "集中・幸運":
            _apply_focus_luck(lv, stats)
        elif name == "極・ダメージ増強":
            _apply_extreme_dmg(lv, stats)
        elif name == "極・適応力":
            _apply_extreme_adapt(lv, stats)
        elif name == "極・HP凝縮":
            _apply_extreme_hp_condensation(lv, stats)
        elif name == "極・応急処置":
            _apply_extreme_first_aid(lv, stats)
        elif name == "極・絶境守護":
            _apply_extreme_desperate_guard(lv, stats)
        elif name == "極・HP変動":
            _apply_extreme_hp_fluctuation(lv, stats)
        elif name == "極・HP吸収":
            _apply_extreme_hp_absorb(lv, stats)
        elif name == "極・幸運会心":
            _apply_extreme_luck_crit(lv, stats)

    return stats


# ==========================================
#  個別エフェクト適用ロジック
# ==========================================

def _apply_agility_boost(lv: int, s: ModuleStats):
    data = {
        6: (30, 40, "敏捷強化: 物理ダメージ+6%"),
        5: (25, 30, "敏捷強化: 物理ダメージ+3.6%"),
        4: (20, 20, None),
        3: (15, 10, None),
        2: (10,  0, None),
        1: (5,   0, None)
    }
    atk, spd, eff = data.get(lv, (0, 0, None))
    s.phys_atk += atk
    s.agility += spd
    if eff:
        s.effects.append(eff)


def _apply_strength_boost(lv: int, s: ModuleStats):
    data = {
        6: (30, 40, "筋力強化: 物理防御力無視+18.8%"),
        5: (25, 30, "筋力強化: 物理防御力無視+11.5%"),
        4: (20, 20, None),
        3: (15, 10, None),
        2: (10,  0, None),
        1: (5,   0, None)
    }
    atk, str_val, eff = data.get(lv, (0, 0, None))
    s.phys_atk += atk
    s.strength += str_val
    if eff:
        s.effects.append(eff)


def _apply_intellect_boost(lv: int, s: ModuleStats):
    data = {
        6: (30, 40, "知力強化: 魔法ダメージ+6%"),
        5: (25, 30, "知力強化: 魔法ダメージ+3.6%"),
        4: (20, 20, None),
        3: (15, 10, None),
        2: (10,  0, None),
        1: (5,   0, None)
    }
    atk, int_val, eff = data.get(lv, (0, 0, None))
    s.mag_atk += atk
    s.intellect += int_val
    if eff:
        s.effects.append(eff)


def _apply_special_dmg_boost(lv: int, s: ModuleStats):
    data = {
        6: (30, 40, "特攻ダメージ強化: 特殊攻撃の属性ダメージ+12%"),
        5: (25, 30, "特攻ダメージ強化: 特殊攻撃の属性ダメージ+7.2%"),
        4: (20, 20, None),
        3: (15, 10, None),
        2: (10,  0, None),
        1: (5,   0, None)
    }
    atk, sia, eff = data.get(lv, (0, 0, None))
    s.pm_atk += atk
    s.sia += sia
    if eff:
        s.effects.append(eff)


def _apply_special_recovery_boost(lv: int, s: ModuleStats):
    data = {
        6: (30, 40, "特攻回復強化: 特殊攻撃回復+12%"),
        5: (25, 30, "特攻回復強化: 特殊攻撃回復+7.2%"),
        4: (20, 20, None),
        3: (15, 10, None),
        2: (10,  0, None),
        1: (5,   0, None)
    }
    atk, int_val, eff = data.get(lv, (0, 0, None))
    s.mag_atk += atk
    s.intellect += int_val
    if eff:
        s.effects.append(eff)


def _apply_mastery_recovery_boost(lv: int, s: ModuleStats):
    data = {
        6: (30, 40, "マスタリー回復強化: マスタリースキル回復+12%"),
        5: (25, 30, "マスタリー回復強化: マスタリースキル回復+7.2%"),
        4: (20, 20, None),
        3: (15, 10, None),
        2: (10,  0, None),
        1: (5,   0, None)
    }
    atk, int_val, eff = data.get(lv, (0, 0, None))
    s.mag_atk += atk
    s.intellect += int_val
    if eff:
        s.effects.append(eff)


def _apply_elite_strike(lv: int, s: ModuleStats):
    data = {
        6: (30, 40, "精鋭打撃: 精鋭以上の対象へのダメージ+6.6%"),
        5: (25, 30, "精鋭打撃: 精鋭以上の対象へのダメージ+3.9%"),
        4: (20, 20, None),
        3: (15, 10, None),
        2: (10,  0, None),
        1: (5,   0, None)
    }
    atk, sia, eff = data.get(lv, (0, 0, None))
    s.pm_atk += atk
    s.sia += sia
    if eff:
        s.effects.append(eff)


def _apply_magic_resist(lv: int, s: ModuleStats):
    data = {
        6: (180, "魔法耐性: 最大HP+4%, 魔法軽減+6%"),
        5: (150, "魔法耐性: 最大HP+3%, 魔法軽減+3.6%"),
        4: (120, "魔法耐性: 最大HP+2%"),
        3: (90,  "魔法耐性: 最大HP+1%"),
        2: (60,  None),
        1: (30,  None)
    }
    endurance, eff = data.get(lv, (0, None))
    s.endurance += endurance
    if eff:
        s.effects.append(eff)


def _apply_phys_resist(lv: int, s: ModuleStats):
    data = {
        6: (480, 20, "物理耐性: 物理軽減+6%"),
        5: (400, 15, "物理耐性: 物理軽減+3.6%"),
        4: (320, 10, None),
        3: (240,  5, None),
        2: (160,  0, None),
        1: (80,   0, None)
    }
    phys_def, attr, eff = data.get(lv, (0, 0, None))
    s.phys_def += phys_def
    s.all_attr += attr
    if eff:
        s.effects.append(eff)


def _apply_focus_atk_speed(lv: int, s: ModuleStats):
    data = {
        6: (50, "集中・攻撃速度: 攻撃速度+6%"),
        5: (40, "集中・攻撃速度: 攻撃速度+3.6%"),
        4: (30, None),
        3: (20, None),
        2: (10, None),
        1: (5,  None)
    }
    atk, eff = data.get(lv, (0, None))
    s.pm_atk += atk
    if eff:
        s.effects.append(eff)


def _apply_focus_cast(lv: int, s: ModuleStats):
    data = {
        6: (50, "集中・詠唱: 詠唱速度+12%"),
        5: (40, "集中・詠唱: 詠唱速度+7.2%"),
        4: (30, None),
        3: (20, None),
        2: (10, None),
        1: (5,  None)
    }
    atk, eff = data.get(lv, (0, None))
    s.pm_atk += atk
    if eff:
        s.effects.append(eff)


def _apply_focus_crit(lv: int, s: ModuleStats):
    data = {
        6: (1800, 80, "集中・会心: 会心ダメージ+12%, 会心回復+12%"),
        5: (1600, 60, "集中・会心: 会心ダメージ+7.1%, 会心回復+7.1%"),
        4: (1200, 40, None),
        3: (900,  20, None),
        2: (600,   0, None),
        1: (300,   0, None)
    }
    hp, attr, eff = data.get(lv, (0, 0, None))
    s.max_hp += hp
    s.all_attr += attr
    if eff:
        s.effects.append(eff)


def _apply_focus_luck(lv: int, s: ModuleStats):
    data = {
        6: (1800, 80, "集中・幸運: 幸運の一撃ダメ倍率+7.8%, 回復倍率+6.2%"),
        5: (1600, 60, "集中・幸運: 幸運の一撃ダメ倍率+4.7%, 回復倍率+3.7%"),
        4: (1200, 40, None),
        3: (900,  20, None),
        2: (600,   0, None),
        1: (300,   0, None)
    }
    hp, attr, eff = data.get(lv, (0, 0, None))
    s.max_hp += hp
    s.all_attr += attr
    if eff:
        s.effects.append(eff)


def _apply_extreme_dmg(lv: int, s: ModuleStats):
    data = {
        6: (60, 80, "極・ダメージ増強: ダメージ時20%の確率で与ダメ+2.75%(最大4スタック/8秒)"),
        5: (50, 60, "極・ダメージ増強: ダメージ時20%の確率で与ダメ+1.65%(最大4スタック/8秒)"),
        4: (40, 40, None),
        3: (30, 20, None),
        2: (20,  0, None),
        1: (10,  0, None)
    }
    atk, sia, eff = data.get(lv, (0, 0, None))
    s.pm_atk += atk
    s.sia += sia
    if eff:
        s.effects.append(eff)


def _apply_extreme_adapt(lv: int, s: ModuleStats):
    data = {
        6: (60, 80, "極・適応力: 開戦後移動速度+30%, 攻撃力+10%(被弾消失/5秒で再獲得)"),
        5: (50, 60, "極・適応力: 開戦後移動速度+18%, 攻撃力+6%(被弾消失/5秒で再獲得)"),
        4: (40, 40, None),
        3: (30, 20, None),
        2: (20,  0, None),
        1: (10,  0, None)
    }
    atk, sia, eff = data.get(lv, (0, 0, None))
    s.pm_atk += atk
    s.sia += sia
    if eff:
        s.effects.append(eff)


def _apply_extreme_hp_condensation(lv: int, s: ModuleStats):
    data = {
        6: (60, 80, "極・HP凝縮: HP50%を上回る超過量5%につき回復量+1.1%"),
        5: (50, 60, "極・HP凝縮: HP50%を上回る超過量5%につき回復量+0.66%"),
        4: (40, 40, None),
        3: (30, 20, None),
        2: (20,  0, None),
        1: (10,  0, None)
    }
    atk, sia, eff = data.get(lv, (0, 0, None))
    s.pm_atk += atk
    s.sia += sia
    if eff:
        s.effects.append(eff)


def _apply_extreme_first_aid(lv: int, s: ModuleStats):
    data = {
        6: (60, 80, "極・応急処置: HP50%未満被弾で5秒間自身と周囲10名HP12%/秒回復"),
        5: (50, 60, "極・応急処置: HP50%未満被弾で5秒間自身と周囲10名HP7.2%/秒回復"),
        4: (40, 40, None),
        3: (30, 20, None),
        2: (20,  0, None),
        1: (10,  0, None)
    }
    atk, int_val, eff = data.get(lv, (0, 0, None))
    s.mag_atk += atk
    s.intellect += int_val
    if eff:
        s.effects.append(eff)


def _apply_extreme_desperate_guard(lv: int, s: ModuleStats):
    data = {
        6: (3600, 80, "極・絶境守護: 軽減3.5% (HP70%未満時20%下がるごとに+4%軽減/最大3)"),
        5: (3000, 60, "極・絶境守護: 軽減2% (HP70%未満時20%下がるごとに+2.4%軽減/最大3)"),
        4: (2400, 40, None),
        3: (1800, 20, None),
        2: (1200,  0, None),
        1: (600,   0, None)
    }
    hp, str_val, eff = data.get(lv, (0, 0, None))
    s.max_hp += hp
    s.strength += str_val
    if eff:
        s.effects.append(eff)


def _apply_extreme_hp_fluctuation(lv: int, s: ModuleStats):
    data = {
        6: (3600, 80, "極・HP変動: 現在HP変化時、最も高いステータス1つが+10%(5秒)"),
        5: (3000, 60, "極・HP変動: 現在HP変化時、最も高いステータス1つが+6%(5秒)"),
        4: (2400, 40, None),
        3: (1800, 20, None),
        2: (1200,  0, None),
        1: (600,   0, None)
    }
    hp, sia, eff = data.get(lv, (0, 0, None))
    s.max_hp += hp
    s.sia += sia
    if eff:
        s.effects.append(eff)


def _apply_extreme_hp_absorb(lv: int, s: ModuleStats):
    data = {
        6: (60, 80, "極・HP吸収: クラススキルでダメージ時、その3%相当のHP回復(最大HP5%迄)"),
        5: (50, 60, "極・HP吸収: クラススキルでダメージ時、その1.8%相当のHP回復(最大HP5%迄)"),
        4: (40, 40, None),
        3: (30, 20, None),
        2: (20,  0, None),
        1: (10,  0, None)
    }
    atk, str_val, eff = data.get(lv, (0, 0, None))
    s.phys_atk += atk
    s.strength += str_val
    if eff:
        s.effects.append(eff)


def _apply_extreme_luck_crit(lv: int, s: ModuleStats):
    data = {
        6: (60, 80, "極・幸運会心: PT全員の会心ダメ+5.2%,幸運ダメ+3.4%(自身は2倍)"),
        5: (50, 60, "極・幸運会心: PT全員の会心ダメ+3.1%,幸運ダメ+2%(自身は2倍)"),
        4: (40, 40, None),
        3: (30, 20, None),
        2: (20,  0, None),
        1: (10,  0, None)
    }
    atk, sia, eff = data.get(lv, (0, 0, None))
    s.pm_atk += atk
    s.sia += sia
    if eff:
        s.effects.append(eff)


# ==========================================
#  CSV検索エンジン (ビームサーチ)
# ==========================================

@dataclass
class CsvModule:
    """CSVから読み込んだ1つのモジュールデータ"""
    id: int = 0
    stats: List[int] = field(default_factory=list)
    heuristic_score: float = 0.0


@dataclass
class CsvSearchCondition:
    """検索条件の1項目"""
    header_name: str = ""
    column_index: int = 0
    priority: SearchPriority = SearchPriority.Normal
    is_extreme: bool = False


@dataclass
class CsvSearchResult:
    """検索結果の1エントリ"""
    score: float = 0.0
    module_ids: List[int] = field(default_factory=list)
    detail_string: str = ""


class CsvSearchEngine:
    """CSV検索エンジン: ビームサーチ (C#版 CsvSearchEngine の移植)"""

    BEAM_WIDTH = 3000
    LV5_BONUS = 150
    LV6_BONUS = 250

    def __init__(self):
        self._modules: List[CsvModule] = []
        self._headers: List[str] = []
        self._is_extreme_col: List[bool] = []
        self._max_stats_values: List[int] = []
        self._module_map: Dict[int, CsvModule] = {}

    def load_csv(self, file_path: str) -> bool:
        """CSVファイルを読み込む"""
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                header_row = next(reader, None)
                if not header_row:
                    return False

                self._headers = header_row[1:]  # ID列を除く
                self._is_extreme_col = ["極" in h for h in self._headers]

                self._modules.clear()
                self._module_map.clear()

                for row in reader:
                    if len(row) < 2:
                        continue
                    try:
                        mod_id = int(row[0])
                    except ValueError:
                        continue

                    stats = []
                    for j in range(1, len(row)):
                        try:
                            stats.append(int(row[j]))
                        except ValueError:
                            stats.append(0)

                    m = CsvModule(id=mod_id, stats=stats)
                    self._modules.append(m)
                    if mod_id not in self._module_map:
                        self._module_map[mod_id] = m

            return True
        except Exception:
            return False

    def get_headers(self) -> List[str]:
        return self._headers

    def get_module_details(self, mod_id: int) -> Optional[Dict[str, int]]:
        """指定IDのモジュールの詳細（効果名: 値）を返す"""
        m = self._module_map.get(mod_id)
        if m is None:
            return None
        result = {}
        for i in range(min(len(self._headers), len(m.stats))):
            if m.stats[i] > 0:
                result[self._headers[i]] = m.stats[i]
        return result

    @staticmethod
    def _get_step_score(val: int) -> float:
        """C#版 GetStepScore"""
        if val >= 20:
            return 31.0
        if val >= 16:
            return 24.5
        if val >= 12:
            return 18.0
        if val >= 8:
            return 11.5
        if val >= 4:
            return 5.0
        if val >= 1:
            return 2.5
        return 0.0

    def search(self, conditions: List[CsvSearchCondition], num_modules: int) -> List[CsvSearchResult]:
        """ビームサーチで最適組み合わせを検索 (同期版)"""
        active_conds = list(conditions)

        # 優先度マップ
        priority_map = [0] * len(self._headers)
        for c in active_conds:
            if c.priority == SearchPriority.Priority:
                priority_map[c.column_index] = 2
            else:
                priority_map[c.column_index] = 1

        # ヒューリスティックスコア計算
        candidates: List[CsvModule] = []
        for m in self._modules:

            h_score = 0.0
            total_stats = 0
            for i in range(min(len(m.stats), len(priority_map))):
                if priority_map[i] > 0:
                    projected_val = m.stats[i] * num_modules
                    base_score = self._get_step_score(projected_val)
                    if i < len(self._is_extreme_col) and self._is_extreme_col[i]:
                        base_score *= 2.0
                    h_score += base_score
                    if priority_map[i] == 2:  # Priority
                        if projected_val >= 20:
                            h_score += self.LV6_BONUS
                        elif projected_val >= 16:
                            h_score += self.LV5_BONUS
                total_stats += m.stats[i]

            h_score += total_stats * num_modules
            m.heuristic_score = h_score
            candidates.append(m)

        candidates.sort(key=lambda x: x.heuristic_score, reverse=True)

        # 枝刈り用最大値テーブル
        if candidates:
            stat_len = len(candidates[0].stats)
            self._max_stats_values = [0] * stat_len
            for m in candidates:
                for i in range(stat_len):
                    if i < len(m.stats) and m.stats[i] > self._max_stats_values[i]:
                        self._max_stats_values[i] = m.stats[i]

        # ビームサーチ
        header_count = len(self._headers)

        class SearchState:
            __slots__ = ['module_indices', 'current_sum', 'current_heuristic_sum', 'last_index']

            def __init__(self):
                self.module_indices: List[int] = []
                self.current_sum: List[int] = [0] * header_count
                self.current_heuristic_sum: float = 0.0
                self.last_index: int = -1

            def clone(self):
                s = SearchState()
                s.module_indices = list(self.module_indices)
                s.current_sum = list(self.current_sum)
                s.current_heuristic_sum = self.current_heuristic_sum
                s.last_index = self.last_index
                return s

        current_states = [SearchState()]

        for step in range(num_modules):
            next_states: List[SearchState] = []
            remaining_steps = num_modules - (step + 1)

            for state in current_states:
                for i in range(state.last_index + 1, len(candidates)):
                    cand = candidates[i]


                    ns = state.clone()
                    ns.module_indices.append(i)
                    ns.last_index = i
                    ns.current_heuristic_sum += cand.heuristic_score

                    for si in range(header_count):
                        if si < len(cand.stats):
                            ns.current_sum[si] += cand.stats[si]
                    next_states.append(ns)

            if len(next_states) > self.BEAM_WIDTH:
                next_states.sort(key=lambda s: s.current_heuristic_sum, reverse=True)
                next_states = next_states[:self.BEAM_WIDTH]

            current_states = next_states
            if not current_states:
                break

        # 最終スコア計算
        results: List[CsvSearchResult] = []

        for state in current_states:
            score = 0.0
            for c in active_conds:
                val = state.current_sum[c.column_index]
                if val > 0:
                    base_score = self._get_step_score(val)
                    if c.column_index < len(self._is_extreme_col) and self._is_extreme_col[c.column_index]:
                        base_score *= 2.0
                    score += base_score
                    if c.priority == SearchPriority.Priority:
                        if val >= 20:
                            score += self.LV6_BONUS
                        elif val >= 16:
                            score += self.LV5_BONUS

            total_all = sum(state.current_sum)
            score += total_all

            detail_parts = []
            for i in range(len(self._headers)):
                if state.current_sum[i] > 0:
                    detail_parts.append(f"{self._headers[i]}:{state.current_sum[i]}")

            result = CsvSearchResult(
                score=score,
                module_ids=[candidates[idx].id for idx in state.module_indices],
                detail_string=" ".join(detail_parts),
            )
            results.append(result)

        results.sort(key=lambda r: r.score, reverse=True)
        return results[:20]

    def search_async(self, conditions: List[CsvSearchCondition], num_modules: int,
                     callback: Callable[[List[CsvSearchResult]], None]):
        """別スレッドで検索を実行し、完了時にcallbackを呼ぶ"""
        def _run():
            res = self.search(conditions, num_modules)
            callback(res)
        t = threading.Thread(target=_run, daemon=True)
        t.start()
