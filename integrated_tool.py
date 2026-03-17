import cv2
import numpy as np
import os
import sys
import json
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk, ImageGrab
import ctypes # Windowsの高DPI設定対策
import mss
import threading
import time
import copy
from module_tab import ModuleTab
import urllib.request
import urllib.error
import subprocess

# ==========================================
#  ★ アプリバージョン ★
# ==========================================
APP_VERSION = "0.3.0"
GITHUB_REPO = "krknt/BPSR-ModuleIntegrater"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

# Windowsの高DPI設定(拡大率)による座標ズレを防ぐおまじない
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ==========================================
#  ★ 日本語変換設定 ★
# ==========================================
TRANSLATION_MAP = {
    "mod_damage_stack": "極・ダメージ増強",
    "mod_life_wave":"極・HP変動",
    "mod_agile": "極・適応力",
    "mod_elite_strike": "精鋭打撃",
    "mod_special_sttack":"特攻ダメージ強化",
    "mod_strength_boost":"筋力強化",
    "mod_agility_boost": "敏捷強化",
    "mod_intellect_boost":"知力強化",
    "mod_attack_spd": "集中・攻撃速度",
    "mod_cast_focus": "集中・詠唱",
    "mod_crit_focus": "集中・会心",
    "mod_luck_focus":"集中・幸運",
    "mod_team_luck_and_crit":"極・幸運会心",
    "mod_first_aid":"極・応急処置",
    "mod_life_condense":"極・HP凝縮",
    "mod_healing_boost":"特攻回復強化",
    "mod_healing_enhance":"マスタリー回復強化",
    "mod_final_protection":"極・絶境守護",
    "mod_life_steal":"極・HP吸収",
    "mod_armor": "物理耐性",
    "mod_resistance":"魔法耐性",
}

def translate_effect(name):
    return TRANSLATION_MAP.get(name, name)

# ==========================================
#  ★ exe化対応用のパス解決関数 ★
# ==========================================
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# ================= 画面キャプチャ用クラス =================

class ScreenSnipper(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.attributes('-fullscreen', True)
        self.attributes('-alpha', 0.3) # 半透明にする
        self.configure(bg='black')
        self.cursor_start = None
        
        self.canvas = tk.Canvas(self, cursor="cross", bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Escape>", lambda e: self.destroy()) # ESCでキャンセル

        self.rect = None

    def on_press(self, event):
        self.cursor_start = (event.x, event.y)
        self.rect = self.canvas.create_rectangle(event.x, event.y, event.x, event.y, outline='red', width=2, fill="white", stipple="gray25")

    def on_drag(self, event):
        if self.cursor_start:
            x, y = self.cursor_start
            self.canvas.coords(self.rect, x, y, event.x, event.y)

    def on_release(self, event):
        if not self.cursor_start: return
        
        x1, y1 = self.cursor_start
        x2, y2 = event.x, event.y
        
        # 座標の正規化
        left, right = sorted([x1, x2])
        top, bottom = sorted([y1, y2])
        
        # ウィンドウを閉じる
        self.destroy()
        self.update() # 画面更新待ち
        
        width = right - left
        height = bottom - top

        # 極小サイズは無視
        if width < 5 or height < 5:
            return

        # スクリーンショット撮影
        try:
            img = ImageGrab.grab(bbox=(left, top, right, bottom))
            # mss用に座標情報を辞書として作成
            monitor_area = {"top": int(top), "left": int(left), "width": int(width), "height": int(height)}
            self.callback(img, monitor_area)
        except Exception as e:
            messagebox.showerror("Error", f"キャプチャに失敗しました: {e}")

# ================= ロジッククラス =================

class GameAnalyzer:
    def __init__(self):
        self.min_icon_width = 10
        self.left_area_percent = 0.1
        self.color_weight = 0.2
        
        self.block_size = 11
        self.shape_threshold = 0.60
        self.digit_threshold = 0.60
        self.template_scale = 1.0
        self.final_threshold = 0.55
        
        self.loaded_image = None
        self.sample_template_size = (50, 50)

    def set_params(self, bs, shape_th, digit_th, scale):
        self.block_size = int(bs) if int(bs) % 2 == 1 else int(bs) + 1
        self.shape_threshold = float(shape_th)
        self.digit_threshold = float(digit_th)
        self.template_scale = float(scale)

    def load_image_from_file(self, path):
        self.loaded_image = cv2.imread(path)
        self._update_sample_size()
        return self.loaded_image is not None

    def load_image_from_memory(self, pil_img):
        open_cv_image = np.array(pil_img) 
        if len(open_cv_image.shape) == 3 and open_cv_image.shape[2] == 3:
            self.loaded_image = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2BGR)
        elif len(open_cv_image.shape) == 3 and open_cv_image.shape[2] == 4:
            self.loaded_image = cv2.cvtColor(open_cv_image, cv2.COLOR_RGBA2BGR)
        else:
            self.loaded_image = open_cv_image
        self._update_sample_size()
        return self.loaded_image is not None

    def _update_sample_size(self):
        folder = resource_path('templates/sub')
        if os.path.exists(folder):
            for f in os.listdir(folder):
                if f.lower().endswith(('.png', '.jpg', '.webp')):
                    img = cv2.imread(os.path.join(folder, f), cv2.IMREAD_UNCHANGED)
                    if img is not None:
                        h, w = img.shape[:2]
                        self.sample_template_size = (w, h)
                        return
        self.sample_template_size = (50, 50)

    def get_preview_image(self):
        if self.loaded_image is None: return None
        img = self.loaded_image.copy()
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        try:
            binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                           cv2.THRESH_BINARY_INV, self.block_size, 3)
        except: return None
        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            is_valid = (x < img.shape[1] * self.left_area_percent and 
                        w > self.min_icon_width and 30 < h)
            color = (0, 255, 0) if is_valid else (0, 0, 255)
            thickness = 2 if is_valid else 1
            cv2.rectangle(img, (x, y), (x + w, y + h), color, thickness)
        return cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    def get_valid_blocks(self):
        if self.loaded_image is None: return []
        gray = cv2.cvtColor(self.loaded_image, cv2.COLOR_BGR2GRAY)
        binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                       cv2.THRESH_BINARY_INV, self.block_size, 3)
        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        raw_blocks = []
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            if (x < self.loaded_image.shape[1] * self.left_area_percent and 
                w > self.min_icon_width and 30 < h):
                raw_blocks.append((x, y, w, h))
        return self._filter_overlapping_blocks(raw_blocks)

    def _filter_overlapping_blocks(self, blocks):
        if not blocks: return []
        blocks.sort(key=lambda b: b[1])
        merged = []
        for b in blocks:
            if not merged:
                merged.append(b)
                continue
            last = merged[-1]
            last_bottom = last[1] + last[3]
            curr_top = b[1]
            if curr_top < last_bottom - 5:
                if abs(last[1] - b[1]) < 20: continue
            merged.append(b)
        return merged

    # --- 解析ロジック ---
    def load_template(self, path):
        img = cv2.imread(path, cv2.IMREAD_UNCHANGED)
        if img is None: return None, None
        if self.template_scale != 1.0:
            h, w = img.shape[:2]
            new_w, new_h = int(w * self.template_scale), int(h * self.template_scale)
            if new_w > 0 and new_h > 0:
                img = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_LINEAR)
        if img.shape[2] == 4:
            bgr = img[:, :, :3]; mask = img[:, :, 3]
            if cv2.countNonZero(mask) == 0: return bgr, None 
            return bgr, mask
        return img, None

    def get_color_score(self, img1, img2, mask=None):
        try:
            hsv1 = cv2.cvtColor(img1, cv2.COLOR_BGR2HSV)
            hsv2 = cv2.cvtColor(img2, cv2.COLOR_BGR2HSV)
            hist1 = cv2.calcHist([hsv1], [0, 1], None, [180, 256], [0, 180, 0, 256])
            hist2 = cv2.calcHist([hsv2], [0, 1], mask, [180, 256], [0, 180, 0, 256])
            cv2.normalize(hist1, hist1, 0, 1, cv2.NORM_MINMAX)
            cv2.normalize(hist2, hist2, 0, 1, cv2.NORM_MINMAX)
            return cv2.compareHist(hist1, hist2, cv2.HISTCMP_CORREL)
        except: return 0.0

    def scan_folder(self, target, folder, is_digit=False):
        detections = []
        if not os.path.exists(folder): return detections
        current_th = self.digit_threshold if is_digit else self.shape_threshold
        for fname in os.listdir(folder):
            if not fname.lower().endswith(('.png','.jpg','.webp')): continue
            tbgr, tmask = self.load_template(os.path.join(folder, fname))
            if tbgr is None: continue
            h, w = tbgr.shape[:2]
            if h > target.shape[0] or w > target.shape[1]: continue
            res = cv2.matchTemplate(target, tbgr, cv2.TM_CCOEFF_NORMED)
            loc = np.where(res >= current_th)
            for pt in zip(*loc[::-1]):
                shape_score = res[pt[1], pt[0]]
                crop = target[pt[1]:pt[1]+h, pt[0]:pt[0]+w]
                if is_digit: final_score = shape_score
                else:
                    c_score = self.get_color_score(crop, tbgr, tmask)
                    final_score = (shape_score * (1 - self.color_weight)) + (c_score * self.color_weight)
                if final_score > current_th:
                    detections.append({'name':os.path.splitext(fname)[0], 'x':int(pt[0]), 'y':int(pt[1]), 'w':w, 'h':h, 'score':float(final_score)})
        return detections

    def remove_duplicates(self, dets):
        if not dets: return []
        dets.sort(key=lambda x: x['score'], reverse=True)
        final = []
        while dets:
            best = dets.pop(0)
            if best['score'] < self.final_threshold: continue
            final.append(best)
            th_x, th_y = best['w']/2, best['h']/2
            dets = [d for d in dets if abs(d['x']-best['x']) > th_x or abs(d['y']-best['y']) > th_y]
        return final

    def resolve_conflicts(self, dets):
        if not dets: return []
        for d in dets:
            name = d['name']
            if name in ['1', 'plus', 'minus', '|']: d['score'] *= 0.85
            if name == '7': d['score'] *= 0.90
            if name == '4': d['score'] *= 0.95
            if name == '8': d['score'] *= 0.95
            if name in ['5', '6', '9']: d['score'] *= 1.05
            if name in ['2', '3']: d['score'] *= 1.02
        dets.sort(key=lambda x: x['score'], reverse=True)
        final = []
        while dets:
            best = dets.pop(0)
            if best['score'] < self.digit_threshold: continue
            final.append(best)
            overlap = best['w'] / 2 
            dets = [d for d in dets if abs((best['x']+best['w']/2) - (d['x']+d['w']/2)) > overlap]
        return final

    def sanitize(self, val):
        clean = "".join(filter(str.isdigit, val))
        if not clean: return val.replace("++","+") if "+" in val else val
        vi = int(clean)
        while vi > 10: 
            vi //= 10
            if vi == 0: break
        return f"+{vi}"

    def analyze_row(self, row_img):
        m_cands = self.scan_folder(row_img, resource_path('templates/main'))
        s_cands = self.scan_folder(row_img, resource_path('templates/sub'))
        d_cands = self.scan_folder(row_img, resource_path('templates/digits'), is_digit=True)
        main = self.remove_duplicates(m_cands)
        subs = self.remove_duplicates(s_cands)
        digits = self.resolve_conflicts(d_cands)
        valid_digits = []
        for d in digits:
            d_cx, d_cy = d['x']+d['w']/2, d['y']+d['h']/2
            if not any(icon['x'] < d_cx < icon['x']+icon['w'] and icon['y'] < d_cy < icon['y']+icon['h'] for icon in subs):
                valid_digits.append(d)
        digits = valid_digits
        subs.sort(key=lambda x: x['x'])
        digits.sort(key=lambda x: x['x'])
        res_stats = []
        for i, icon in enumerate(subs):
            related = []
            ir = icon['x'] + icon['w']
            icy = icon['y'] + icon['h']/2
            limit = subs[i+1]['x'] if i+1 < len(subs) else row_img.shape[1]
            for d in digits:
                if ir-5 < d['x'] < limit and abs(icy - (d['y']+d['h']/2)) < 20:
                    related.append(d)
            if related:
                related.sort(key=lambda x: x['x'])
                raw_v = "".join([d['name'] for d in related]).replace("plus","+")
                res_stats.append({"type": icon['name'], "value": self.sanitize(raw_v)})
        return {"main_icon": main[0]['name'] if main else None, "stats": res_stats}

# ================= GUIクラス =================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("BPSRモジュールキャプチャ")
        try:
            self.root.iconbitmap(resource_path('icon.ico'))
        except Exception:
            pass
        self.analyzer = GameAnalyzer()
        self.img_path = "captured" 
        self.current_results = []
        self.stock_data = []
        self.display_scale = 1.0

        # ★ リアルタイム処理用変数
        self.capture_area = None
        self.is_running_realtime = False
        self.last_result_hash = "" 
        self.result_history_buffer = [] # 重複チェック用バッファ(直近2回分)

        # ★ ユーザーオプション
        self.zero_pad_empty = tk.BooleanVar(value=True)
        self.show_advanced_params = False # コラプスの初期状態

        # Menu bar
        self.menu_bar = tk.Menu(self.root)
        self.option_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.option_menu.add_command(label="出力設定", command=self.open_settings)
        self.option_menu.add_separator()
        self.option_menu.add_command(label="アップデート確認", command=self.check_for_update)
        self.menu_bar.add_cascade(label="⚙ 設定", menu=self.option_menu)
        self.root.config(menu=self.menu_bar)

        # ★ ttk.Notebook でタブ化
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # --- キャプチャタブ ---
        self.capture_tab = tk.Frame(self.notebook)
        self.notebook.add(self.capture_tab, text="キャプチャ")

        # --- モジュールタブ ---
        self.module_tab_frame = tk.Frame(self.notebook)
        self.notebook.add(self.module_tab_frame, text="モジュール")
        self.module_tab = ModuleTab(self.module_tab_frame)

        # Layout (キャプチャタブ内)
        self.ctrl_frame = tk.Frame(self.capture_tab, width=300, padx=10, pady=10, bg="#f0f0f0")
        self.ctrl_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.ctrl_frame.pack_propagate(False)

        self.right_frame = tk.Frame(self.capture_tab, width=350, bg="#ffffff")
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.right_frame.pack_propagate(False)

        self.view_frame = tk.Frame(self.capture_tab, bg="gray")
        self.view_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(self.view_frame, bg="gray", cursor="arrow")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Guide
        self.guide_rect = None
        self.drag_data = {"x": 0, "y": 0, "item": None}
        self.canvas.bind("<ButtonPress-1>", self.on_drag_start)
        self.canvas.bind("<B1-Motion>", self.on_drag_motion)
        self.guide_rect = self.canvas.create_rectangle(-100,-100,0,0, outline="cyan", width=3, tags="guide")

        # --- Control Panel ---
        tk.Label(self.ctrl_frame, text="操作パネル", bg="#f0f0f0", font=("Meiryo UI", 12, "bold")).pack(pady=(0,10))
        
        # ボタンエリア
        btn_frame = tk.Frame(self.ctrl_frame, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(btn_frame, text="画像を開く", command=self.open_file, bg="white", width=10).pack(side=tk.LEFT, padx=(0,2))
        tk.Button(btn_frame, text="画面撮影", command=self.start_capture, bg="#ffeebb", width=10).pack(side=tk.LEFT, padx=2)
        
        # ★ 自動監視ボタン
        self.btn_realtime = tk.Button(btn_frame, text="監視 開始", command=self.toggle_realtime, bg="#ffccff", width=10)
        self.btn_realtime.pack(side=tk.LEFT, padx=2)
        
        # 分割設定
        lbl_split = self._create_label("分割 (緑枠) の調整")
        lbl_split.pack(pady=(15,5), anchor="w")
        
        self.block_frame, self.s_block = self._create_slider_stepper(self.ctrl_frame, "ブロックサイズ:", 3, 31, 11, res=2)
        self.block_frame.pack(fill=tk.X, pady=(0, 5))

        # 認識設定
        lbl_recog = self._create_label("認識精度の調整")
        lbl_recog.pack(pady=(15,5), anchor="w")
        
        self.scale_frame, self.s_scale = self._create_slider_stepper(self.ctrl_frame, "サイズ倍率:", 0.5, 2.0, 1.0, res=0.05)
        self.scale_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.advanced_widgets = []
        
        # 詳細設定コラプスのトグルボタン
        self.btn_collapse = tk.Button(self.ctrl_frame, text="▼ 詳細パラメータを開く", command=self.toggle_advanced_ui, bg="#e0e0e0", relief=tk.FLAT)
        self.btn_collapse.pack(fill=tk.X, pady=(15, 5))

        # 詳細パラメータを格納する専用のフレーム（ボタンの直下に配置）
        self.advanced_frame = tk.Frame(self.ctrl_frame, bg="#f0f0f0")
        
        self.shape_frame, self.s_shape = self._create_slider_stepper(self.advanced_frame, "アイコン精度:", 0.1, 1.0, 0.60, res=0.05)
        self.advanced_widgets.append(self.shape_frame)
        
        self.digit_frame, self.s_digit = self._create_slider_stepper(self.advanced_frame, "数字精度:", 0.1, 1.0, 0.60, res=0.05) 
        self.advanced_widgets.append(self.digit_frame)
        
        self.show_guide = tk.BooleanVar(value=True)
        self.chk_guide = self._create_checkbox(self.advanced_frame, "青枠表示", self.show_guide, self.toggle_guide)
        self.advanced_widgets.append(self.chk_guide)
        
        # 初期状態は隠すため pack() せず、必要時に toggle_advanced_ui で pack する
        self.update_advanced_ui()

        self.is_topmost = tk.BooleanVar(value=False)
        self.topmost_frame = self._create_checkbox(self.ctrl_frame, "最前面に固定", self.is_topmost, self.toggle_topmost)
        self.topmost_frame.pack(fill=tk.X, pady=5)

        tk.Button(self.ctrl_frame, text="解析実行 (手動)", command=self.run_analysis, 
                  bg="lightblue", height=2, font=("Meiryo UI", 11, "bold")).pack(fill=tk.X, pady=20)

        for s in [self.s_block, self.s_scale, self.s_shape, self.s_digit]:
            s.config(command=self.update_preview)

        # 設定の読み込み
        self._load_config()

        # 終了時に設定を保存
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # --- Right Panel ---
        self.curr_frame = tk.LabelFrame(self.right_frame, text="現在の解析結果 (ダブルクリック修正)", bg="white")
        self.curr_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        tree_scroll = tk.Scrollbar(self.curr_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree = ttk.Treeview(self.curr_frame, columns=("Val"), yscrollcommand=tree_scroll.set)
        self.tree.heading("#0", text="項目名")
        self.tree.heading("Val", text="値")
        self.tree.column("#0", width=160)
        self.tree.column("Val", width=80)
        self.tree.pack(fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.tree.yview)
        self.tree.bind("<Double-1>", self.on_double_click)

        tk.Button(self.curr_frame, text="↓ リストに追加 ↓", command=self.add_to_stock, 
                  bg="#ffdddd", font=("Meiryo UI", 10, "bold")).pack(fill=tk.X, pady=5)

        self.stock_frame = tk.LabelFrame(self.right_frame, text="ストック済み", bg="white")
        self.stock_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.stock_listbox = tk.Listbox(self.stock_frame)
        self.stock_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        stock_scroll = tk.Scrollbar(self.stock_frame, command=self.stock_listbox.yview)
        stock_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.stock_listbox.config(yscrollcommand=stock_scroll.set)

        tk.Button(self.right_frame, text="まとめて保存 (CSV/JSON)", command=self.export_all, 
                  bg="lightgreen", height=2, font=("Meiryo UI", 11, "bold")).pack(fill=tk.X, padx=5, pady=10)

    # --- Menu Options Logic ---
    def open_settings(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("出力設定")
        settings_win.geometry("300x220")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()

        # メインウィンドウの中央に配置
        settings_win.update_idletasks()
        sw = settings_win.winfo_width()
        sh = settings_win.winfo_height()
        rx = self.root.winfo_rootx()
        ry = self.root.winfo_rooty()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        x = rx + (rw - sw) // 2
        y = ry + (rh - sh) // 2
        settings_win.geometry(f"+{x}+{y}")

        frame = tk.Frame(settings_win, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="エクスポート設定", font=("Meiryo UI", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))

        chk_frame = tk.Frame(frame)
        chk_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(chk_frame, text="空の値を0埋めで出力 (CSV/JSON)", font=("Meiryo UI", 9)).pack(side=tk.LEFT)
        tk.Checkbutton(chk_frame, variable=self.zero_pad_empty).pack(side=tk.RIGHT)

        tk.Button(frame, text="設定フォルダを開く", command=self._open_config_dir, font=("Meiryo UI", 9)).pack(fill=tk.X, pady=(10, 0))

        tk.Button(frame, text="閉じる", command=settings_win.destroy, width=10).pack(pady=(15, 0))

    def toggle_advanced_ui(self):
        self.show_advanced_params = not self.show_advanced_params
        if self.show_advanced_params:
            self.btn_collapse.config(text="▲ 詳細パラメータを閉じる")
        else:
            self.btn_collapse.config(text="▼ 詳細パラメータを開く")
        self.update_advanced_ui()

    def update_advanced_ui(self):
        show = self.show_advanced_params
        if show:
            for w in self.advanced_widgets:
                w.pack(fill=tk.X, pady=(2, 2))
                        
            # ボタンの直下にフレームごと表示する
            self.advanced_frame.pack(fill=tk.X, after=self.btn_collapse)
        else:
            self.advanced_frame.pack_forget()

    # --- GUI Helper Methods ---
    def _create_label(self, text):
        lbl = tk.Label(self.ctrl_frame, text=f"{text}", bg="#f0f0f0", font=("Meiryo UI", 9, "bold"))
        return lbl
        
    def _create_slider_stepper(self, parent_frame, text, min_v, max_v, def_v, res=1):
        outer = tk.Frame(parent_frame, bg="#f0f0f0")
        
        var = tk.DoubleVar(value=def_v) if isinstance(def_v, float) else tk.IntVar(value=def_v)

        val_lbl = tk.Label(outer, text=f"{def_v:.2f}" if isinstance(def_v, float) else f"{def_v}", bg="#f0f0f0", font=("Meiryo UI", 9, "bold"), width=4, anchor="e")

        def update_lbl(*args):
            v = var.get()
            val_lbl.config(text=f"{v:.2f}" if isinstance(def_v, float) else f"{v}")
        var.trace_add("write", update_lbl)

        def adjust(delta):
            v = var.get() + delta
            if v < min_v: v = min_v
            if v > max_v: v = max_v
            
            if isinstance(def_v, float):
                v_f = float(round(float(v / res)) * res)
                v_f = float(f"{v_f:.4f}")
                var.set(v_f) # type: ignore
            else:
                v_i = int(round(float(v / res)) * res)
                var.set(v_i) # type: ignore
                
            self.update_preview()

        # 配置: [名称][スペース][-][スライダー][+][数値]
        tk.Label(outer, text=text, bg="#f0f0f0", font=("Meiryo UI", 9)).pack(side=tk.LEFT)

        # 右側から逆順に配置
        val_lbl.pack(side=tk.RIGHT, padx=(2, 0))

        btn_plus = tk.Button(outer, text="＋", command=lambda: adjust(res), font=("Meiryo UI", 8, "bold"), bg="#ffffff", bd=1, padx=4, pady=0)
        btn_plus.pack(side=tk.RIGHT, padx=0)

        scale = tk.Scale(outer, from_=min_v, to=max_v, orient=tk.HORIZONTAL, resolution=res, 
                         showvalue=False, variable=var, bg="#f0f0f0", sliderlength=20, length=80, width=15)
        scale.bind("<ButtonRelease-1>", lambda e: self.update_preview())
        scale.pack(side=tk.RIGHT, padx=0)

        btn_minus = tk.Button(outer, text="－", command=lambda: adjust(-res), font=("Meiryo UI", 8, "bold"), bg="#ffffff", bd=1, padx=4, pady=0)
        btn_minus.pack(side=tk.RIGHT, padx=0)

        scale.get = var.get
        return outer, scale

    def _create_checkbox(self, parent_frame, text, variable, command=None):
        """[名称][いい感じのスペース][チェックボックス] のレイアウトを返す"""
        frame = tk.Frame(parent_frame, bg="#f0f0f0")
        tk.Label(frame, text=text, bg="#f0f0f0", font=("Meiryo UI", 9)).pack(side=tk.LEFT)
        chk = tk.Checkbutton(frame, variable=variable, command=command, bg="#f0f0f0")
        chk.pack(side=tk.RIGHT)
        return frame

    def _add_label(self, text):
        lbl = tk.Label(self.ctrl_frame, text=f"{text}", bg="#f0f0f0", font=("Meiryo UI", 9, "bold"))
        lbl.pack(pady=(15,5), anchor="w")
        return lbl
        
    def _add_slider(self, label, min_v, max_v, def_v, res=1):
        tk.Label(self.ctrl_frame, text=label, bg="#f0f0f0", font=("Meiryo UI", 9)).pack(anchor="w")
        scale = tk.Scale(self.ctrl_frame, from_=min_v, to=max_v, orient=tk.HORIZONTAL, resolution=res, bg="#f0f0f0")
        scale.set(def_v)
        scale.pack(fill=tk.X, pady=(0, 5))
        return scale

    # --- Capture Logic ---
    def start_capture(self):
        self.root.iconify()
        ScreenSnipper(self.root, self.finish_capture)

    def finish_capture(self, img, monitor=None):
        self.root.deiconify()
        if img:
            self.img_path = "ScreenCapture"
            self.capture_area = monitor 
            if self.analyzer.load_image_from_memory(img):
                self.update_preview()
                self.clear_tree()
                self.current_results = []
                self.root.title(f"Target: {self.img_path} (Ready for Monitor)")
                self.show_guide.set(True)
                self.toggle_guide()

    # --- Real-time Logic ---
    def toggle_realtime(self):
        if self.is_running_realtime:
            self.is_running_realtime = False
            self.btn_realtime.config(text="監視 開始", bg="#ffccff", fg="black")
            messagebox.showinfo("停止", "自動監視を停止しました。")
        else:
            if not self.capture_area:
                messagebox.showwarning("エラー", "先に「画面撮影」で監視エリアを指定してください。")
                return
            
            self.is_running_realtime = True
            self.result_history_buffer = [] # バッファリセット
            self.btn_realtime.config(text="■ 停止", bg="red", fg="white")
            self.root.title("監視中... (Esc等で操作可能)")
            
            thread = threading.Thread(target=self.realtime_loop)
            thread.daemon = True 
            thread.start()

    def realtime_loop(self):
        with mss.mss() as sct:
            while self.is_running_realtime:
                start_time = time.time()

                try:
                    sct_img = sct.grab(self.capture_area)
                    img_np = np.array(sct_img)
                    img_bgr = cv2.cvtColor(img_np, cv2.COLOR_BGRA2BGR)
                except Exception as e:
                    print(f"Capture Error: {e}")
                    break

                self.analyzer.loaded_image = img_bgr
                blocks = self.analyzer.get_valid_blocks()
                
                if blocks:
                    row_results = []
                    for i, (bx, by, bw, bh) in enumerate(blocks):
                        y1, y2 = max(0, by-2), min(img_bgr.shape[0], by+bh+2)
                        row_img = img_bgr[y1:y2, max(0, bx-5):img_bgr.shape[1]]
                        res = self.analyzer.analyze_row(row_img)
                        if res['main_icon'] or res['stats']:
                            row_results.append(res)
                    
                    # 単純な全体ハッシュ比較ではなく、データが有る場合に進む
                    if row_results:
                        self.root.after(0, self.add_realtime_data, f"Auto_{time.strftime('%H:%M:%S')}", row_results, img_bgr)

                elapsed = time.time() - start_time
                wait = max(1.5 - elapsed, 0.1) 
                time.sleep(wait)

    def _rows_are_equal(self, row_a, row_b):
        """ 行単位での同一性判定 """
        if row_a.get('main_icon') != row_b.get('main_icon'): return False
        stats_a = row_a.get('stats', [])
        stats_b = row_b.get('stats', [])
        if len(stats_a) != len(stats_b): return False
        return json.dumps(stats_a, sort_keys=True) == json.dumps(stats_b, sort_keys=True)

    def add_realtime_data(self, fname, results, img_bgr):
        """ GUI更新 & 重複排除 """
        unique_results = []
        
        # 今回検出された各行について、過去のバッファ内にあるかチェック
        for new_row in results:
            is_duplicate = False
            for past_frame_results in self.result_history_buffer:
                for past_row in past_frame_results:
                    if self._rows_are_equal(new_row, past_row):
                        is_duplicate = True
                        break
                if is_duplicate: break
            
            if not is_duplicate:
                unique_results.append(new_row)

        # バッファ更新（重複判定用に、フィルタ前の全データを保存）
        self.result_history_buffer.insert(0, results)
        if len(self.result_history_buffer) > 2:
            self.result_history_buffer.pop()

        if not unique_results:
            return # 新規データなし

        # ストック追加
        self.stock_data.append({"filename": fname, "data": unique_results})
        self.stock_listbox.insert(tk.END, f"★ {fname} (新規: {len(unique_results)} 行)")
        self.stock_listbox.yview(tk.END)
        
        # プレビュー更新
        self.update_preview()
        
        # ツリービュー表示
        self.current_results = unique_results
        self.clear_tree()
        for i, res in enumerate(unique_results):
            main_txt = f"Row {i}: {res['main_icon']}" if res['main_icon'] else f"Row {i}"
            row_id = self.tree.insert("", "end", text=main_txt, values=(""))
            for stat in res["stats"]:
                jp_name = translate_effect(stat["type"])
                self.tree.insert(row_id, "end", text=jp_name, values=(stat["value"]))
            self.tree.item(row_id, open=True)

    # --- Drag & Guide Logic ---
    def on_drag_start(self, event):
        if not self.show_guide.get(): return
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        self.drag_data["item"] = self.guide_rect

    def on_drag_motion(self, event):
        if not self.show_guide.get(): return
        delta_x = event.x - self.drag_data["x"]
        delta_y = event.y - self.drag_data["y"]
        self.canvas.move(self.guide_rect, delta_x, delta_y)
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y

    def toggle_guide(self):
        state = 'normal' if self.show_guide.get() else 'hidden'
        self.canvas.itemconfigure(self.guide_rect, state=state)
        if state == 'normal':
            cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
            self.update_guide_box_size(cw/2, ch/2)
            self.canvas.tag_raise(self.guide_rect)

    def update_guide_box_size(self, cx=None, cy=None):
        scale = self.s_scale.get()
        base_w, base_h = self.analyzer.sample_template_size
        
        disp_w = base_w * scale * self.display_scale
        disp_h = base_h * scale * self.display_scale
        
        coords = self.canvas.coords(self.guide_rect)
        if cx is None:
            if coords:
                cx = (coords[0] + coords[2]) / 2
                cy = (coords[1] + coords[3]) / 2
            else:
                cx, cy = 100, 100
        
        self.canvas.coords(self.guide_rect, cx - disp_w/2, cy - disp_h/2, cx + disp_w/2, cy + disp_h/2)

    # --- Core Functions ---
    def open_file(self):
        if self.is_running_realtime:
            self.toggle_realtime() 

        f = filedialog.askopenfilename(filetypes=[("Image", "*.png;*.jpg;*.webp")])
        if f:
            self.img_path = f
            self.capture_area = None 
            if self.analyzer.load_image_from_file(f):
                self.update_preview()
                self.clear_tree()
                self.current_results = []
                self.root.title(f"Target: {os.path.basename(f)}")
                self.show_guide.set(True)
                self.toggle_guide()

    def update_preview(self, event=None):
        if self.analyzer.loaded_image is None: return

        bs = int(self.s_block.get())
        if bs % 2 == 0: bs += 1
        
        self.analyzer.set_params(
            bs, 
            self.s_shape.get(), 
            self.s_digit.get(), 
            self.s_scale.get()
        )
        
        prev = self.analyzer.get_preview_image()
        if prev is None: return
        
        h, w = prev.shape[:2]
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        
        self.display_scale = 1.0
        if cw > 1 and ch > 1:
            self.display_scale = min(cw/w, ch/h)
            
        new_w, new_h = int(w * self.display_scale), int(h * self.display_scale)
        prev = cv2.resize(prev, (new_w, new_h))
        
        self.pil = ImageTk.PhotoImage(image=Image.fromarray(prev))
        self.canvas.delete("img_tag")
        self.canvas.create_image(cw//2, ch//2, anchor=tk.CENTER, image=self.pil, tags="img_tag")
        self.canvas.tag_lower("img_tag")
        
        self.update_guide_box_size()

    def run_analysis(self):
        if self.analyzer.loaded_image is None: return
        self.update_preview()
        blocks = self.analyzer.get_valid_blocks()
        if not blocks:
            messagebox.showwarning("警告", "行が見つかりません")
            return
        
        self.current_results = []
        self.clear_tree()
        img = self.analyzer.loaded_image
        
        for i, (bx, by, bw, bh) in enumerate(blocks):
            y1, y2 = max(0, by-2), min(img.shape[0], by+bh+2)
            row_img = img[y1:y2, max(0, bx-5):img.shape[1]]
            res = self.analyzer.analyze_row(row_img)
            res["id"] = i
            self.current_results.append(res)
            main_txt = f"Row {i}: {res['main_icon']}" if res['main_icon'] else f"Row {i}"
            row_id = self.tree.insert("", "end", text=main_txt, values=(""))
            for stat in res["stats"]:
                jp_name = translate_effect(stat["type"])
                self.tree.insert(row_id, "end", text=jp_name, values=(stat["value"]))
            self.tree.item(row_id, open=True)

    def clear_tree(self):
        for item in self.tree.get_children(): self.tree.delete(item)

    def on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        column = self.tree.identify_column(event.x)
        if column != "#1": return 
        item_id = self.tree.identify_row(event.y)
        parent_id = self.tree.parent(item_id)
        if not parent_id: return 
        bbox = self.tree.bbox(item_id, column)
        entry = tk.Entry(self.tree)
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry.insert(0, self.tree.item(item_id, "values")[0])
        entry.select_range(0, tk.END)
        entry.focus()
        def save_edit(e=None):
            new_val = entry.get()
            self.tree.set(item_id, column="Val", value=new_val)
            r_idx = self.tree.index(parent_id)
            s_idx = self.tree.index(item_id)
            self.current_results[r_idx]["stats"][s_idx]["value"] = new_val
            entry.destroy()
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

    def add_to_stock(self):
        if not self.current_results: return
        fname = os.path.basename(self.img_path)
        self.stock_data.append({"filename": fname, "data": self.current_results})
        self.stock_listbox.insert(tk.END, f"{len(self.stock_data)}. {fname} ({len(self.current_results)} 行)")
        self.clear_tree()
        self.current_results = []
        messagebox.showinfo("追加", "リストに追加しました。")

    def export_all(self):
        if not self.stock_data:
            messagebox.showwarning("警告", "ストックが空です")
            return
            
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv"), ("JSON", "*.json")])
        if not path: return

        all_effect_types = set()
        for item in self.stock_data:
            for row in item['data']:
                for stat in row['stats']:
                    all_effect_types.add(stat['type'])
        
        header_jp = ["ID"] + list(TRANSLATION_MAP.values())

        # エクスポート用にデータを複製し、nullや空文字をオプションに従って変換
        export_data = copy.deepcopy(self.stock_data)
        zero_pad = self.zero_pad_empty.get()
        for item in export_data:
            for row in item['data']:
                for stat in row['stats']:
                    val = stat.get('value', "0" if zero_pad else "")
                    if val is None or str(val).strip().lower() in ["", "null", "none", "+"]:
                        stat['value'] = "0" if zero_pad else ""

        if path.endswith(".json"):
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=4, ensure_ascii=False)
        else:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(header_jp)
                
                global_id = 1
                for item in export_data:
                    for row in item['data']:
                        current_stats = {stat['type']: stat['value'] for stat in row['stats']}
                        csv_row = [global_id]
                        for effect in TRANSLATION_MAP.keys():
                            default_val = "0" if zero_pad else ""
                            val = str(current_stats.get(effect, default_val))
                            if val.strip().lower() in ["", "null", "none", "+"]:
                                val = default_val
                            val = val.replace("+","")
                            csv_row.append(val)

                        writer.writerow(csv_row)
                        global_id += 1
                        
        messagebox.showinfo("完了", "保存しました。")

# ★★★ 追加コード: 最前面固定切り替え ★★★
    def toggle_topmost(self):
        self.root.attributes('-topmost', self.is_topmost.get())

# ★★★ アップデート機能 ★★★
    def _cleanup_old_exe(self):
        """起動時に前回の更新で残った .old ファイルを削除する"""
        try:
            old_path = os.path.join(self._get_config_dir(), "previous_version.exe.old")
            if os.path.exists(old_path):
                os.remove(old_path)
        except Exception:
            pass

    def check_for_update(self):
        """メニューから呼び出されるアップデート確認。非同期で実行する。"""
        threading.Thread(target=self._check_for_update_thread, daemon=True).start()

    def _check_for_update_thread(self):
        """バックグラウンドで GitHub API を呼び出して更新を確認する"""
        try:
            req = urllib.request.Request(GITHUB_API_URL, headers={"Accept": "application/vnd.github.v3+json"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            latest_tag = data.get("tag_name", "").lstrip("v")
            if not latest_tag:
                self.root.after(0, lambda: messagebox.showwarning("アップデート", "リリース情報を取得できませんでした。"))
                return

            if self._compare_versions(latest_tag, APP_VERSION) > 0:
                # 新しいバージョンがある
                download_url = None
                for asset in data.get("assets", []):
                    if asset.get("name", "").endswith(".exe"):
                        download_url = asset.get("browser_download_url")
                        break

                if download_url:
                    self.root.after(0, lambda: self._prompt_update(latest_tag, download_url))
                else:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "アップデート",
                        f"新しいバージョン v{latest_tag} がありますが\n"
                        f"ダウンロード可能な .exe が見つかりませんでした。\n"
                        f"GitHub ページを確認してください。"))
            else:
                self.root.after(0, lambda: messagebox.showinfo(
                    "アップデート",
                    f"現在のバージョン v{APP_VERSION} は最新です。"))

        except urllib.error.URLError:
            self.root.after(0, lambda: messagebox.showerror("アップデート", "ネットワークに接続できませんでした。"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("アップデート", f"エラーが発生しました:\n{e}"))

    @staticmethod
    def _compare_versions(v1, v2):
        """バージョン文字列を比較する。v1 > v2 なら正, v1 == v2 なら0, v1 < v2 なら負"""
        parts1 = [int(x) for x in v1.split(".")]
        parts2 = [int(x) for x in v2.split(".")]
        for a, b in zip(parts1, parts2):
            if a != b:
                return a - b
        return len(parts1) - len(parts2)

    def _prompt_update(self, tag_name, download_url):
        """アップデートダイアログを表示する"""
        result = messagebox.askyesno(
            "アップデート",
            f"新しいバージョン v{tag_name} が利用可能です。\n"
            f"現在: v{APP_VERSION}\n\n"
            f"ダウンロードして更新しますか？")
        if result:
            threading.Thread(target=self._do_update, args=(download_url, tag_name), daemon=True).start()

    def _do_update(self, download_url, tag_name):
        """.exe をダウンロードして置き換える"""
        try:
            exe_path = sys.executable
            new_path = exe_path + ".new"
            old_path = os.path.join(self._get_config_dir(), "previous_version.exe.old")

            # ダウンロード
            self.root.after(0, lambda: self.root.title(f"BPSR - ダウンロード中..."))
            urllib.request.urlretrieve(download_url, new_path)

            # 現在の exe を config フォルダの .old にリネーム
            # 注意: 同一ドライブ内であれば os.rename が高速
            if os.path.exists(old_path):
                os.remove(old_path)
            os.rename(exe_path, old_path)

            # 新しい exe を元のパスにリネーム
            os.rename(new_path, exe_path)

            self.root.after(0, lambda: self._update_complete(tag_name, exe_path))

        except Exception as e:
            # ロールバック: .new があれば削除、.old があれば戻す
            try:
                if os.path.exists(new_path) and not os.path.exists(exe_path):
                    if os.path.exists(old_path):
                        os.rename(old_path, exe_path)
                if os.path.exists(new_path):
                    os.remove(new_path)
            except Exception:
                pass
            self.root.after(0, lambda: messagebox.showerror("アップデート失敗", f"更新に失敗しました:\n{e}"))
            self.root.after(0, lambda: self.root.title("BPSR - Module Integrater"))

    def _update_complete(self, tag_name, exe_path):
        """アップデート完了後の処理"""
        self.root.title("BPSR - Module Integrater")
        result = messagebox.askyesno(
            "アップデート完了",
            f"v{tag_name} へのアップデートが完了しました。\n"
            f"アプリを再起動しますか？")
        if result:
            subprocess.Popen([exe_path])
            self.root.destroy()

# ★★★ 設定の保存・読み込み ★★★
    CONFIG_FILE = "config.json"

    def _get_config_dir(self):
        """設定フォルダのパスを返す (%APPDATA%\BPSR-ModuleIntegrater\)"""
        appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
        config_dir = os.path.join(appdata, "BPSR-ModuleIntegrater")
        os.makedirs(config_dir, exist_ok=True)
        return config_dir

    def _get_config_path(self):
        """設定ファイルのパスを返す"""
        return os.path.join(self._get_config_dir(), self.CONFIG_FILE)

    def _save_config(self):
        """現在のパラメータを JSON に保存する"""
        config = {
            "block_size": self.s_block.get(),
            "size_scale": self.s_scale.get(),
            "shape_threshold": self.s_shape.get(),
            "digit_threshold": self.s_digit.get(),
            "show_guide": self.show_guide.get(),
            "is_topmost": self.is_topmost.get(),
            "zero_pad_empty": self.zero_pad_empty.get(),
            "show_advanced_params": self.show_advanced_params,
        }
        try:
            with open(self._get_config_path(), "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def _load_config(self):
        """保存済みのパラメータを JSON から読み込む"""
        path = self._get_config_path()
        if not os.path.exists(path):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                config = json.load(f)

            # スライダーの値を復元
            if "block_size" in config:
                self.s_block.set(config["block_size"])
            if "size_scale" in config:
                self.s_scale.set(config["size_scale"])
            if "shape_threshold" in config:
                self.s_shape.set(config["shape_threshold"])
            if "digit_threshold" in config:
                self.s_digit.set(config["digit_threshold"])

            # チェックボックスの値を復元
            if "show_guide" in config:
                self.show_guide.set(config["show_guide"])
                self.toggle_guide()
            if "is_topmost" in config:
                self.is_topmost.set(config["is_topmost"])
                self.toggle_topmost()
            if "zero_pad_empty" in config:
                self.zero_pad_empty.set(config["zero_pad_empty"])

            # コラプスの状態を復元
            if "show_advanced_params" in config and config["show_advanced_params"]:
                self.show_advanced_params = True
                self.btn_collapse.config(text="▲ 詳細パラメータを閉じる")
                self.update_advanced_ui()

        except Exception:
            pass

    def _on_close(self):
        """ウィンドウ閉じる時に設定を保存して終了する"""
        self._save_config()
        self.root.destroy()

    def _open_config_dir(self):
        """設定ファイルが保存されているフォルダをエクスプローラーで開く"""
        path = os.path.dirname(self._get_config_path())
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showerror("エラー", "設定フォルダが見つかりません。")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1210x850")
    root.attributes('-topmost', False)
    app = App(root)
    app._cleanup_old_exe()
    root.mainloop()