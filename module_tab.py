import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Optional

from module_models import (
    EFFECT_TYPES, MODULE_EFFECT_NAMES,
    SearchPriority,
    ModuleOption, ModuleItem, ModuleEffectStatus, ModuleStats,
    CsvSearchCondition, CsvSearchResult, CsvSearchEngine,
    update_module_effects, calculate_module_stats,
)

# ==========================================
#  カラー定数
# ==========================================
COL_PURPLE = "#8E44AD"
COL_GREEN = "#2E7D32"
COL_GREEN_BG = "#F1F8E9"
COL_GREEN_BORDER = "#558B2F"
COL_GREEN_BTN = "#27AE60"
COL_BLUE = "#3498DB"
COL_RED = "#E74C3C"
COL_BG = "#FAFAFA"
COL_BORDER = "#DDD"
CARD_WIDTH = 280
CARD_PAD = 8


class ModuleTab:
    """モジュールタブ全体のUI + ロジック"""

    MAX_EQUIPPED = 4

    def __init__(self, parent_frame: tk.Frame):
        self.parent = parent_frame

        # データ
        self.modules: List[ModuleItem] = []
        self.effect_summary: List[ModuleEffectStatus] = []
        self.link_total = 0
        self.csv_engine = CsvSearchEngine()
        self.csv_conditions: List[CsvSearchCondition] = []
        self.csv_results: List[CsvSearchResult] = []

        # 効果サマリー初期化
        for name in MODULE_EFFECT_NAMES:
            self.effect_summary.append(ModuleEffectStatus(name=name))

        # 初期モジュール (4枠)
        for i in range(1, 5):
            item = ModuleItem(
                name=f"モジュール {i}",
                is_equipped=True,
                on_update=self._on_module_changed,
                on_equip_changed=self._on_equip_changed,
            )
            self.modules.append(item)

        self._build_ui()
        self._refresh_all()

    # ==========================================
    #  UI構築
    # ==========================================
    def _build_ui(self):
        # メインスクロール
        canvas = tk.Canvas(self.parent, bg=COL_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.parent, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg=COL_BG)

        # --- スクロール領域を正しく計算する関数 ---
        def _update_scrollregion(event=None):
            bbox = canvas.bbox("all")
            if bbox:
                # 描画内容の高さと、キャンバス(ウィンドウ)の高さを比較
                canvas_h = canvas.winfo_height()
                # 内容が少ない時はキャンバスの高さを最小値として設定し、スライドを防ぐ
                scroll_h = max(bbox[3], canvas_h)
                canvas.configure(scrollregion=(0, 0, bbox[2], scroll_h))

        self.scroll_frame.bind("<Configure>", _update_scrollregion)
        
        self._canvas_win = canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # スクロールフレームをキャンバス幅に合わせる
        def _on_canvas_configure(event):
            canvas.itemconfigure(self._canvas_win, width=event.width)
            _update_scrollregion() # ウィンドウリサイズ時も再計算
            
        canvas.bind("<Configure>", _on_canvas_configure)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # マウスホイール対応
        def _on_mousewheel(event):
            # ▼追加: イベントの発生源がリストボックス等の場合は、全体のスクロールを無視する
            if isinstance(event.widget, (tk.Listbox, tk.Text, tk.Canvas)):
                # Listboxのスクロールバーなどを操作している時は何もしない
                if event.widget != canvas:
                    return

            # 中身がキャンバスより大きい時（本当にスクロールが必要な時）だけ許可
            if self.scroll_frame.winfo_reqheight() > canvas.winfo_height():
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self._canvas = canvas

        # --- セクション1: CSV検索ツール ---
        self._build_csv_section()

        # --- セクション2: リンク効果サマリー ---
        self._build_effect_summary_section()

        # --- セクション3: モジュール効果一覧 ---
        self._build_effect_list_section()

        # --- セクション4: 装備モジュール ---
        self._build_modules_section()

    # ─────────── CSV検索セクション ───────────
    def _build_csv_section(self):
        grp = tk.LabelFrame(self.scroll_frame, text="CSVデータベース検索ツール",
                            bg="white", fg=COL_PURPLE, font=("Meiryo UI", 10, "bold"),
                            padx=10, pady=10)
        grp.pack(fill=tk.X, padx=10, pady=(10, 5))

        tk.Label(grp, text="CSVファイルを読み込んで、最適な組み合わせを検索します",
                 fg="#666", bg="white", font=("Meiryo UI", 9)).pack(anchor="w", pady=(0, 10))

        # ボタン行
        btn_row = tk.Frame(grp, bg="white")
        btn_row.pack(fill=tk.X, pady=(0, 10))

        tk.Button(btn_row, text="📂 CSV読込", command=self._load_csv,
                  bg=COL_PURPLE, fg="white", font=("Meiryo UI", 9, "bold"),
                  padx=15, pady=4).pack(side=tk.LEFT)

        tk.Label(btn_row, text="組み合わせ数:", bg="white",
                 font=("Meiryo UI", 9)).pack(side=tk.LEFT, padx=(20, 5))
        self.combo_count_var = tk.IntVar(value=4)
        combo = ttk.Combobox(btn_row, textvariable=self.combo_count_var,
                             values=[2, 3, 4, 5], width=3, state="readonly")
        combo.pack(side=tk.LEFT)

        tk.Button(btn_row, text="🔍 検索実行", command=self._run_csv_search,
                  bg=COL_PURPLE, fg="white", font=("Meiryo UI", 9, "bold"),
                  padx=20, pady=4).pack(side=tk.RIGHT)

        # 条件パネル
        self.cond_frame = tk.Frame(grp, bg="white", bd=1, relief=tk.SOLID)
        self.cond_frame.pack(fill=tk.X, pady=(0, 10))
        self.cond_inner = tk.Frame(self.cond_frame, bg="white")
        self.cond_inner.pack(fill=tk.X)

        # 検索結果
        tk.Label(grp, text="検索結果 (Top 20)", bg="white",
                 font=("Meiryo UI", 9, "bold")).pack(anchor="w", pady=(0, 5))

        result_frame = tk.Frame(grp, bg="white")
        result_frame.pack(fill=tk.X)

        result_scroll = tk.Scrollbar(result_frame)
        result_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.result_listbox = tk.Listbox(result_frame, height=8,
                                          yscrollcommand=result_scroll.set,
                                          font=("Meiryo UI", 9))
        self.result_listbox.pack(fill=tk.X, expand=True)
        result_scroll.config(command=self.result_listbox.yview)

        btn_apply_row = tk.Frame(grp, bg="white")
        btn_apply_row.pack(fill=tk.X, pady=(5, 0))
        tk.Button(btn_apply_row, text="選択した構成を装備", command=self._apply_selected_result,
                  bg=COL_GREEN_BTN, fg="white", font=("Meiryo UI", 9, "bold"),
                  padx=10, pady=3).pack(side=tk.RIGHT)

        # ステータスラベル
        self.status_var = tk.StringVar(value="")
        tk.Label(grp, textvariable=self.status_var, bg="white", fg="#666",
                 font=("Meiryo UI", 8)).pack(anchor="w")

    # ─────────── リンク効果サマリーセクション ───────────
    def _build_effect_summary_section(self):
        self.summary_frame = tk.Frame(self.scroll_frame, bg=COL_GREEN_BG,
                                       bd=1, relief=tk.SOLID, padx=10, pady=10)
        self.summary_frame.pack(fill=tk.X, padx=10, pady=5)

        # 左: 総リンク効果
        left = tk.Frame(self.summary_frame, bg=COL_GREEN_BG)
        left.pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(left, text="総リンク効果", bg=COL_GREEN_BG, fg=COL_GREEN_BORDER,
                 font=("Meiryo UI", 10)).pack()
        self.link_total_var = tk.StringVar(value="0")
        tk.Label(left, textvariable=self.link_total_var, bg=COL_GREEN_BG,
                 fg=COL_GREEN, font=("Meiryo UI", 24, "bold")).pack()
        self.link_agility_var = tk.StringVar(value="敏捷 +0")
        tk.Label(left, textvariable=self.link_agility_var, bg=COL_GREEN,
                 fg="white", font=("Meiryo UI", 9, "bold"), padx=8, pady=2).pack()

        # 右: 効果カード群
        self.effect_cards_frame = tk.Frame(self.summary_frame, bg=COL_GREEN_BG)
        self.effect_cards_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # ─────────── 効果テキスト一覧セクション ───────────
    def _build_effect_list_section(self):
        self.effect_list_frame = tk.LabelFrame(self.scroll_frame, text="モジュール効果一覧",
                                                bg="white", font=("Meiryo UI", 10, "bold"),
                                                padx=10, pady=5)
        self.effect_list_frame.pack(fill=tk.X, padx=10, pady=5)

        # 基礎ステータス表示
        self.base_stats_var = tk.StringVar(value="")
        self.base_stats_label = tk.Label(self.effect_list_frame, textvariable=self.base_stats_var,
                                          bg="#E8F5E9", fg=COL_GREEN,
                                          font=("Meiryo UI", 9, "bold"),
                                          anchor="w", padx=5, pady=3)
        self.base_stats_label.pack(fill=tk.X, pady=(0, 5))

        # ユニーク効果表示
        self.effects_text = tk.Text(self.effect_list_frame, height=6, wrap=tk.WORD,
                                     font=("Meiryo UI", 9), state=tk.DISABLED,
                                     bg="white", bd=0)
        self.effects_text.pack(fill=tk.X)

    # ─────────── モジュールカードセクション ───────────
    def _build_modules_section(self):
        # ヘッダー
        header_frame = tk.Frame(self.scroll_frame, bg=COL_BG)
        header_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        tk.Label(header_frame, text="現在の装備モジュール", bg=COL_BG,
                 font=("Meiryo UI", 14, "bold")).pack(side=tk.LEFT)

        self.equipped_count_var = tk.StringVar(value="装備中: 4")
        tk.Label(header_frame, textvariable=self.equipped_count_var, bg=COL_BG,
                 fg=COL_GREEN, font=("Meiryo UI", 12, "bold")).pack(side=tk.LEFT, padx=15)

        tk.Button(header_frame, text="+ 空スロット追加", command=self._add_module,
                  bg=COL_BLUE, fg="white", font=("Meiryo UI", 9, "bold"),
                  padx=15, pady=3).pack(side=tk.RIGHT)

        # カード一覧フレーム (WrapPanel風)
        self.cards_frame = tk.Frame(self.scroll_frame, bg=COL_BG)
        self.cards_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        self._card_widgets: List[tk.Frame] = []
        self.cards_frame.bind("<Configure>", lambda e: self._reflow_cards())

    # ==========================================
    #  UI更新
    # ==========================================
    def _refresh_all(self):
        """データからUI全体を再描画"""
        self.link_total = update_module_effects(self.modules, self.effect_summary)
        stats = calculate_module_stats(self.effect_summary, self.link_total)

        # リンク効果
        self.link_total_var.set(str(self.link_total))
        self.link_agility_var.set(f"敏捷 +{self.link_total * 2}")

        # 装備数
        eq_count = sum(1 for m in self.modules if m.is_equipped)
        self.equipped_count_var.set(f"装備中: {eq_count}")

        # 効果カード
        self._refresh_effect_cards()

        # 基礎ステータス
        self._refresh_stats_display(stats)

        # モジュールカード
        self._refresh_module_cards()

    def _refresh_effect_cards(self):
        """効果サマリーカードを再描画"""
        for w in self.effect_cards_frame.winfo_children():
            w.destroy()

        row_frame = None
        col = 0
        max_cols = 4

        for status in self.effect_summary:
            if not status.is_visible:
                continue

            if col % max_cols == 0:
                row_frame = tk.Frame(self.effect_cards_frame, bg=COL_GREEN_BG)
                row_frame.pack(fill=tk.X, pady=1)
            col += 1

            card = tk.Frame(row_frame, bg="white", bd=1, relief=tk.SOLID, padx=6, pady=4)
            card.pack(side=tk.LEFT, padx=3, pady=2)

            # 上段: 名前 + Lv
            top = tk.Frame(card, bg="white")
            top.pack(fill=tk.X)
            tk.Label(top, text=status.name, bg="white", fg="#333",
                     font=("Meiryo UI", 9, "bold"), width=12, anchor="w").pack(side=tk.LEFT)
            lv_frame = tk.Frame(top, bg="#E8F5E9", padx=4)
            lv_frame.pack(side=tk.RIGHT)
            tk.Label(lv_frame, text=f"Lv.{status.level}", bg="#E8F5E9",
                     fg=COL_GREEN, font=("Meiryo UI", 8, "bold")).pack()

            # 下段: プログレスバー + 数値
            bottom = tk.Frame(card, bg="white")
            bottom.pack(fill=tk.X, pady=(2, 0))
            prog = ttk.Progressbar(bottom, value=status.current_total,
                                    maximum=status.next_threshold, length=100)
            prog.pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Label(bottom, text=f"{status.current_total}/{status.next_threshold}",
                     bg="white", fg="#666", font=("Meiryo UI", 8)).pack(side=tk.RIGHT, padx=(4, 0))

    def _refresh_stats_display(self, stats: ModuleStats):
        """基礎ステータスと効果テキストを更新"""
        parts = []
        if stats.phys_atk > 0:
            parts.append(f"物理攻撃力 +{stats.phys_atk}")
        if stats.mag_atk > 0:
            parts.append(f"魔法攻撃力 +{stats.mag_atk}")
        if stats.pm_atk > 0:
            parts.append(f"物理/魔法攻撃力 +{stats.pm_atk}")
        if stats.strength > 0:
            parts.append(f"筋力 +{stats.strength}")
        if stats.intellect > 0:
            parts.append(f"知力 +{stats.intellect}")
        if stats.agility > 0:
            parts.append(f"敏捷 +{stats.agility}")
        if stats.sia > 0:
            parts.append(f"筋力/知力/敏捷 +{stats.sia}")
        if stats.max_hp > 0:
            parts.append(f"最大HP +{stats.max_hp}")
        if stats.endurance > 0:
            parts.append(f"耐久力 +{stats.endurance}")
        if stats.phys_def > 0:
            parts.append(f"物理防御力 +{stats.phys_def}")
        if stats.all_attr > 0:
            parts.append(f"全属性強度 +{stats.all_attr}")

        if parts:
            self.base_stats_var.set("[基礎ステータス合計] " + ", ".join(parts))
            self.base_stats_label.pack(fill=tk.X, pady=(0, 5))
        else:
            self.base_stats_var.set("")
            self.base_stats_label.pack_forget()

        # ユニーク効果テキスト
        self.effects_text.config(state=tk.NORMAL)
        self.effects_text.delete("1.0", tk.END)
        if stats.effects:
            self.effects_text.insert(tk.END, "\n".join(f"  ▸ {e}" for e in stats.effects))
            self.effects_text.config(height=min(len(stats.effects) + 1, 10))
        else:
            self.effects_text.insert(tk.END, "  (効果なし)")
            self.effects_text.config(height=2)
        self.effects_text.config(state=tk.DISABLED)

    def _refresh_module_cards(self):
        """モジュールカードを再描画"""
        for w in self.cards_frame.winfo_children():
            w.destroy()
        self._card_widgets.clear()

        for idx, mod in enumerate(self.modules):
            card = self._build_module_card(self.cards_frame, mod, idx)
            self._card_widgets.append(card)

        self.cards_frame.after_idle(self._reflow_cards)

    def _reflow_cards(self):
        """ウィンドウ幅に応じてカードを折り返し配置する"""
        container_w = self.cards_frame.winfo_width()
        if container_w <= 1:
            return

        x, y = CARD_PAD, CARD_PAD
        row_height = 0
        for card in self._card_widgets:
            card.update_idletasks()
            ch = card.winfo_reqheight()
            if x + CARD_WIDTH + CARD_PAD > container_w and x > CARD_PAD:
                x = CARD_PAD
                y += row_height + CARD_PAD
                row_height = 0
            card.place(x=x, y=y, width=CARD_WIDTH)
            row_height = max(row_height, ch)
            x += CARD_WIDTH + CARD_PAD

        total_h = y + row_height + CARD_PAD
        self.cards_frame.config(height=max(total_h, 10))

    def _build_module_card(self, parent: tk.Frame, mod: ModuleItem, idx: int) -> tk.Frame:
        """1つのモジュールカードを構築して返す"""
        card = tk.Frame(parent, bg="white", bd=1, relief=tk.SOLID, padx=10, pady=8)

        # 上段: チェック + 名前 + 削除
        top = tk.Frame(card, bg="white")
        top.pack(fill=tk.X, pady=(0, 8))

        equip_var = tk.BooleanVar(value=mod.is_equipped)

        def on_toggle():
            mod.set_equipped(equip_var.get())

        chk = tk.Checkbutton(top, variable=equip_var, command=on_toggle, bg="white")
        chk.pack(side=tk.LEFT)

        name_var = tk.StringVar(value=mod.name)

        def on_name_change(*args):
            mod.name = name_var.get()

        name_entry = tk.Entry(top, textvariable=name_var, font=("Meiryo UI", 10, "bold"),
                              bd=0, bg="white", width=14)
        name_entry.pack(side=tk.LEFT, padx=5)
        name_var.trace_add("write", on_name_change)

        if mod.is_equipped:
            tk.Label(top, text="装備中", bg="white", fg=COL_GREEN,
                     font=("Meiryo UI", 8, "bold")).pack(side=tk.LEFT)

        def on_delete():
            self._delete_module(idx)
        tk.Button(top, text="✕", command=on_delete, bg=COL_RED, fg="white",
                  font=("Meiryo UI", 8, "bold"), width=2, height=1, bd=0).pack(side=tk.RIGHT)

        # 下段: オプション×3
        for opt_idx, opt in enumerate(mod.options):
            opt_frame = tk.Frame(card, bg="white")
            opt_frame.pack(fill=tk.X, pady=1)

            type_var = tk.StringVar(value=opt.selected_type)

            def make_type_cb(o, v):
                def cb(*args):
                    o.set_type(v.get())
                return cb

            type_combo = ttk.Combobox(opt_frame, textvariable=type_var,
                                       values=EFFECT_TYPES, width=16, state="readonly",
                                       font=("Meiryo UI", 8))
            type_combo.pack(side=tk.LEFT)
            type_var.trace_add("write", make_type_cb(opt, type_var))

            val_var = tk.IntVar(value=opt.value)

            def make_val_cb(o, v):
                def cb(*args):
                    try:
                        o.set_value(v.get())
                    except tk.TclError:
                        pass
                return cb

            val_spin = tk.Spinbox(opt_frame, from_=0, to=10, textvariable=val_var,
                                   width=3, font=("Meiryo UI", 9),
                                   justify=tk.CENTER)
            val_spin.pack(side=tk.LEFT, padx=(3, 0))
            val_var.trace_add("write", make_val_cb(opt, val_var))

            tk.Label(opt_frame, text="pt", bg="white", fg="#888",
                     font=("Meiryo UI", 8)).pack(side=tk.LEFT)

        return card

    # ==========================================
    #  CSV検索条件UI
    # ==========================================
    def _refresh_conditions_ui(self):
        """条件パネルをデータから再構築"""
        for w in self.cond_inner.winfo_children():
            w.destroy()

        if not self.csv_conditions:
            tk.Label(self.cond_inner, text="CSVを読み込んでください", bg="white",
                     fg="#999", font=("Meiryo UI", 9)).pack(pady=10)
            return

        # 3列グリッド
        max_cols = 3
        for i, cond in enumerate(self.csv_conditions):
            row = i // max_cols
            col = i % max_cols

            def make_toggle(c, btn_ref):
                def toggle():
                    if c.priority == SearchPriority.Normal:
                        c.priority = SearchPriority.Priority
                        btn_ref[0].config(bg=COL_PURPLE, fg="white",
                                          text=f"★ {c.header_name}")
                    else:
                        c.priority = SearchPriority.Normal
                        btn_ref[0].config(bg="#E0E0E0", fg="#333",
                                          text=c.header_name)
                return toggle

            is_priority = cond.priority == SearchPriority.Priority
            btn_text = f"★ {cond.header_name}" if is_priority else cond.header_name
            btn_bg = COL_PURPLE if is_priority else "#E0E0E0"
            btn_fg = "white" if is_priority else "#333"

            btn_holder = [None]  # mutable ref for closure
            btn = tk.Button(self.cond_inner, text=btn_text,
                            bg=btn_bg, fg=btn_fg,
                            font=("Meiryo UI", 9), relief=tk.FLAT,
                            padx=8, pady=4, anchor="w")
            btn_holder[0] = btn
            btn.config(command=make_toggle(cond, btn_holder))
            btn.grid(row=row, column=col, sticky="ew", padx=2, pady=2)

        for c in range(max_cols):
            self.cond_inner.columnconfigure(c, weight=1)

    # ==========================================
    #  CSV検索結果UI
    # ==========================================
    def _refresh_results_ui(self):
        """結果リストを更新"""
        self.result_listbox.delete(0, tk.END)
        for r in self.csv_results:
            ids_str = ", ".join(str(i) for i in r.module_ids)
            self.result_listbox.insert(tk.END,
                                        f"Score: {r.score:.1f}  |  IDs: [{ids_str}]  |  {r.detail_string}")

    # ==========================================
    #  イベントハンドラ
    # ==========================================
    def _on_module_changed(self):
        """モジュールのオプション変更時"""
        self._refresh_all()

    def _on_equip_changed(self, item: ModuleItem):
        """装備ON/OFF変更時"""
        eq_count = sum(1 for m in self.modules if m.is_equipped)
        if eq_count > self.MAX_EQUIPPED:
            item.is_equipped = False
            messagebox.showwarning("制限", f"同時に装備できるモジュールは最大{self.MAX_EQUIPPED}つです。")
            return
        self._refresh_all()

    def _add_module(self):
        """空スロットを追加"""
        next_num = len(self.modules) + 1
        item = ModuleItem(
            name=f"モジュール {next_num}",
            is_equipped=sum(1 for m in self.modules if m.is_equipped) < self.MAX_EQUIPPED,
            on_update=self._on_module_changed,
            on_equip_changed=self._on_equip_changed,
        )
        self.modules.append(item)
        self._refresh_all()

    def _delete_module(self, idx: int):
        """モジュールを削除"""
        if idx < 0 or idx >= len(self.modules):
            return
        name = self.modules[idx].name
        if messagebox.askyesno("確認", f"{name} を削除しますか？"):
            self.modules.pop(idx)
            self._refresh_all()

    def _load_csv(self):
        """CSVファイルを読み込む"""
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not path:
            return
        if self.csv_engine.load_csv(path):
            self.csv_conditions.clear()
            headers = self.csv_engine.get_headers()
            for i, h in enumerate(headers):
                self.csv_conditions.append(CsvSearchCondition(
                    header_name=h,
                    column_index=i,
                    is_extreme="極" in h,
                    priority=SearchPriority.Normal,
                ))
            self._refresh_conditions_ui()
            self.status_var.set("CSVを読み込みました。条件を設定してください。")
        else:
            messagebox.showerror("エラー", "CSVの読み込みに失敗しました。")

    def _run_csv_search(self):
        """CSV検索を実行"""
        if not self.csv_conditions:
            messagebox.showwarning("注意", "先にCSVを読み込んでください。")
            return

        self.status_var.set("検索中... (処理には時間がかかる場合があります)")
        self.parent.update_idletasks()

        num = self.combo_count_var.get()

        def on_complete(results: List[CsvSearchResult]):
            self.csv_results = results
            self.parent.after(0, self._on_search_complete)

        self.csv_engine.search_async(self.csv_conditions, num, on_complete)

    def _on_search_complete(self):
        """検索完了コールバック(GUIスレッドで実行)"""
        self._refresh_results_ui()
        if self.csv_results:
            self.status_var.set(f"{len(self.csv_results)} 件見つかりました。")
        else:
            self.status_var.set("条件に合う組み合わせが見つかりませんでした。")

    def _apply_selected_result(self):
        """選択した検索結果をモジュールに反映"""
        sel = self.result_listbox.curselection()
        if not sel:
            messagebox.showwarning("選択", "結果を選択してください。")
            return
        idx = sel[0]
        if idx >= len(self.csv_results):
            return

        result = self.csv_results[idx]

        # 既存モジュールをクリアして結果を反映
        self.modules.clear()
        for mod_id in result.module_ids:
            details = self.csv_engine.get_module_details(mod_id)
            if details is None:
                continue
            item = ModuleItem(
                name=f"ID:{mod_id} (CSV)",
                is_equipped=True,
                on_update=self._on_module_changed,
                on_equip_changed=self._on_equip_changed,
            )
            opt_idx = 0
            for eff_name, eff_val in details.items():
                if opt_idx >= 3:
                    break
                item.options[opt_idx].selected_type = eff_name
                item.options[opt_idx].value = eff_val
                opt_idx += 1
            self.modules.append(item)

        self._refresh_all()
        self.status_var.set("検索結果の構成を装備しました！")
