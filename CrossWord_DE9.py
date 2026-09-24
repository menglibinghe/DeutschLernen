import pandas as pd
import re
import random
import os
from reportlab.pdfgen import canvas             #创建一个 PDF 文档对象，创建画布
from reportlab.lib.pagesizes import A4          #导入A4页面尺寸常量
from reportlab.pdfbase import pdfmetrics        #负责管理PDF的字体系统，使用非标准字体（如中文字体）时，必须通过这个模块来注册
from reportlab.pdfbase.ttfonts import TTFont    #用于加载和注册TrueType 字体文件 
from reportlab.lib import colors


# ----------------------------------------------
# 版本说明：
# 单词填空时预显的字母按照单词包含字母数量比例随机设置，而不是全局字母随机分布
# 预先字母比例随机化波动±5%，首字母必选，防连字机制，首字母->交叉点=稀有字母->其他的高低权重，权重大的字母预显的概率也会更大
# 包含空格的可分单词，空格按照～正常显示，并且～不算作交叉点构建
# 德语字母ß可以正常实现和交叉布局
# 优化布局，减少因为首字母和交叉字母必须显示导致连续字母显示出现
# --------------------------------------------------


GermansWordFile = "_单词260514.xlsx"
# ==========================================
# 用户自定义配置区
# ==========================================
START_ROW = 375      # Excel第2行对应序号1,336
END_ROW = 396     # 建议范围 15-25，358
ITERATIONS = 150   # 迭代次数
HINT_RATE = 0.5    # 预填字母比例，建议不高于0.5 
# ==========================================
print('当前excel文件为: 单词260514.xlsx')
try:
    START_ROW = int((input('当前请输入起始行号(>=0):  ').strip()))
    END_ROW = int((input('当前请输入终止行号(<起始+24): ').strip()))
    tmp = input('预填字母比例(0.3～0.6, 默认0.5):  ').strip()
    HINT_RATE = 0.5 if tmp =='' else float(tmp)
    print(f'当前预填字母比例为： {HINT_RATE}')
except:
    print('输入的数字格式不正确')
    raise


#FONT_PATH = 'C:/Windows/Fonts/msyh.ttc'
FONT_PATH = 'C:/Windows/Fonts/msyh.ttc'

if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont('msyh', FONT_PATH))  #加载到 reportlab 的字体系统中，使其可以用于 PDF 绘制
    FONT_NAME = 'msyh'
#    FONT_NAME = 'msyh'
else:
    FONT_NAME = 'Helvetica'

#(超级)清洗特殊字符，用于单词列表显示：剔除音标，保留所有标点符号，并将换行符转为 ||
def super_clean_text(text): 
    if pd.isna(text) or text == "": return ""           #isna()函数检查text是否为空（如 None、NaN、NaT）
    text = str(text)                                    #类型强制转换，确保 text 是字符串类型
    
    # 核心修改 1: 在剔除不可见字符前，先将换行符替换为 || 
    text = text.replace('\n', ' || ').replace('\r', ' ')
    
    # 剔除音标内容 [音标]
 #  text = re.sub(r'\[[a-zA-Zεɐ̯ˈɪ].*?\]', '', text)     #匹配方括号内第一个字符及其之后的所有内容，全部去除
    text = text.replace('\xa0', ' ').replace('\u200b', '')  #\xa0是不换行空格(&nbsp;)，\u200b 是零宽空格
    
    # 核心修改 2: 扩展白名单，保留所有标点符号 [! ? . , : ; " ' ( ) + - / =]
    # 使用 ^ 表示取反，即：除了白名单内的字符，全部删掉
    # 这里我们保留了德语常用标点字符集
    safe_pattern = re.compile(r'[^\u4e00-\u9fa5a-zA-Z0-9äöüÄÖÜßẞβ \!\?\.\,\:\;\"\'\(\)\+\-\/\=\|]\[\]')   ##\u4e00-\u9fa5：匹配所有中文字符（Unicode 范围）#safe_pattern = re.compile(r'[^\u4e00-\u9fa5a-zA-Z0-9äöüÄÖÜß \!\?\.\,\:\;\"\'\(\)\+\-\/\=\|]')
    #safe_pattern.sub('', text)删除所有不在白名单中的字符，.split()按任意空白字符（空格、制表符、换行符等）分割字符串，" ".join(...)用单个空格将列表中的元素连接起来,.strip()删除字符串首尾的空白字符   
    return " ".join(safe_pattern.sub('', text).split()).strip() 

#清洗掉音标及特殊字符，转为大写，用于棋盘填空
def clean_for_grid(word):                       #将单词清洗成纯字母形式，作为填词的语料
    if pd.isna(word) or word == "": return ""   
    w = re.sub(r'\[\d+\]', '', str(word))       #[ \d+\]'：匹配方括号包裹的数字，例如 [1]、[123]，然后去除
    w = re.sub(r'[ßẞ]', 'BBB', w)               #先把ß转换为BBB，然后单词转换为大写字母后，再转换为ß，避免与单个或两个BB字母的单词混淆
    # 修改：在保留字符的正则中加入空格 " "
    w = re.sub(r'[^a-zA-ZäöüÄÖÜßẞβ ]', '', w)   
    w = w.strip().upper()
    w = re.sub('BBB', 'ß', w)
    # 修改：将空格替换为波浪号 ~
    w = w.replace(' ', '~')
    return w


class CrosswordEngine:                          #size为棋盘尺寸，grid为棋盘的网格，palced_words为已放置的大写单词，原始单词+释义 
    def __init__(self, size=32):                #构造函数
        self.size = size
        self.grid = [['#' for _ in range(size)] for _ in range(size)]   #嵌套生成[['#', ..., '#'],...,['#', ..., '#']]
        self.placed_words = []                  #清空已放置的单词列表
        # 【修改之处】新增数据成员：用于记录和回传最终筛选出的最优方案的预显字母坐标集合
        self.hint_coords = set()

    #_表示私有成员函数
    def _clear(self):
        self.grid = [['#' for _ in range(self.size)] for _ in range(self.size)]
        self.placed_words = []
        # 【修改之处】清空记录的预填坐标
        self.hint_coords = set()

    def _can_place(self, word_upper, x, y, direction):  #检查棋盘有无空间可以放置单词
        if x < 0 or y < 0 or (direction == 'H' and x + len(word_upper) > self.size) or \
           (direction == 'V' and y + len(word_upper) > self.size):
            return False
        
        # 计算当前字母的坐标
        has_overlap = False
        for i in range(len(word_upper)): 
            #查询当前字母对应坐标                   
            curr_x = x + (i if direction == 'H' else 0)     #如果()里面条件为假则i=0，否则不变
            curr_y = y + (0 if direction == 'H' else i)
            char_in_grid = self.grid[curr_y][curr_x]        #当前坐标位置已存在的字符
            
            # 冲突检测
            if char_in_grid != '#' and char_in_grid != word_upper[i]:
                return False
            # 修改：只有当重叠的不是 ~ 时，才认为是一个有效的交叉点
            if char_in_grid != '#' and char_in_grid != '~': 
                has_overlap = True            


            #邻居(粘连)检测：如果当前位置的邻居(上下、左右或者旁边)存在字母，则返回false表示不可放置
            if char_in_grid == '#':     #如果当前坐标位置位置没有字符
                checks = []             #检查列表初始为空
                if direction == 'H':    #水平方向上，检查当前坐标的上、下邻居元组
                    checks = [(curr_x, curr_y-1), (curr_x, curr_y+1)]
                    if i == 0: checks.append((curr_x-1, curr_y))    #如果待放置的字母为单词首字母，还要添加左侧邻居
                    if i == len(word_upper)-1: checks.append((curr_x+1, curr_y))    #如果待放置的字母为单词末字母，还要添加右侧邻居
                else:
                    checks = [(curr_x-1, curr_y), (curr_x+1, curr_y)]
                    if i == 0: checks.append((curr_x, curr_y-1))
                    if i == len(word_upper)-1: checks.append((curr_x, curr_y+1))
                
                #检查邻居列表，如有邻居不为空，则认为不可放置
                for cx, cy in checks:
                    if 0 <= cx < self.size and 0 <= cy < self.size:
                        if self.grid[cy][cx] != '#': return False     #棋盘矩阵里面行坐标为cy，列坐标为cx
        #检查已放置单词列表，如果不为空，要求当前单词必须与其存在交叉现象
        return has_overlap if self.placed_words else True   #如果条件为真，返回as_overlap，否则返回True
    
    #执行单词放置
    def _place(self, word_upper, full_row_data, x, y, direction):
        for i, char in enumerate(word_upper):
            gx, gy = (x + i, y) if direction == 'H' else (x, y + i)
            self.grid[gy][gx] = char            #在当前位置上开始连续放置每个字母，棋盘矩阵里面行坐标为cy，列坐标为cx        
        
        # 记录放置的单词信息，类型:字典列表
        self.placed_words.append({
            'word_upper': word_upper,           #放置大写单词
            'full_info': full_row_data,         #Excel 表格中当前处理的那一行的所有原始数据（经过清洗后）
            'x': x, 'y': y,                     #单词放置在棋盘中的首字母的位置坐标
            'dir': direction
        })

    #检查并放置单词
    def _try_add(self, word_upper, full_row_data):
        # 1. 如果是第一个单词，直接水平方向放在中间
        if not self.placed_words:
            self._place(word_upper, full_row_data, self.size // 3, self.size // 3, 'H')
            return True
        possible_moves = []
        for i, char in enumerate(word_upper):                                   # 遍历待放置的单词字母，i为待放置字母的位置游标
            for item in self.placed_words:                                      # 遍历已放置的每个单词
                for j, pchar in enumerate(item['word_upper']):                  # 遍历已放置单词的每个字母，j为匹配到的已放置单词的字母的位置游标
                    if char == pchar:                                           #寻找字母交点：如果字母相同
                        new_dir = 'V' if item['dir'] == 'H' else 'H'    #横、纵向轮流放置
                        # 计算新单词的起始坐标 (nx, ny)
                        nx = item['x'] + j if item['dir'] == 'H' else item['x'] - i     #已有单词横向排列时item['x'] + j为重叠字母位置横坐标，也是新单词首字母的横坐标；否则为item['x'] - i 
                        ny = item['y'] - i if item['dir'] == 'H' else item['y'] + j     #已有单词横向排列时item['y'] - i为重叠字母位置横坐标，也是新单词首字母的横坐标；否则为item['y'] + j 
                        if self._can_place(word_upper, nx, ny, new_dir):
                            possible_moves.append((nx, ny, new_dir))
        if possible_moves:
            nx, ny, ndir = random.choice(possible_moves)        #随机选择匹配的现有单词或者同一个单词后多个重叠字母对应的新单词应放置的起点坐标
            self._place(word_upper, full_row_data, nx, ny, ndir)
            return True
        return False

    # 【修改之处】全新进化升级版本的 generate_best_layout 函数
    # 功能：通过多次随机尝试，在满足最大覆盖所输入单词列表的前提下，选择连续字母出现最少的优胜方案
    def generate_best_layout(self, raw_data, iterations=100, hint_rate=0.5, session_seed=42):
        best_state = None
        max_count = -1
        min_streak_violations = float('inf')  # 记录全局最少连字违规数

        for attempt in range(iterations):
            self._clear()
            temp_list = raw_data[:]     #暂存原始数据元组列表
            random.shuffle(temp_list)   ## 打乱顺序并按长度排序（长单词优先通常更难放，先放）
            temp_list.sort(key=lambda x: len(x[0]), reverse=True)   #lambda x: len(x[0])创建内嵌函数，sort按照key指定的内嵌函数输出，按照降序排列
            for wu, fd in temp_list: 
                self._try_add(wu, fd)
            
            current_placed_count = len(self.placed_words)   # 判断当前这一轮迭代生成的布局
            
            # 优胜劣汰准则 1：如果当前放入的单词数比历史最高还少，直接淘汰，不浪费算力
            if current_placed_count < max_count:
                continue
                
            # ==========================================================
            # 质检前置：在当前网格中完全模拟最终 PDF 的提示字选取过程
            # ==========================================================
            current_hint_coords = set()
            grid_usage = {}
            for item in self.placed_words:
                for i in range(len(item['word_upper'])):
                    gx = item['x'] + (i if item['dir'] == 'H' else 0)
                    gy = item['y'] + (0 if item['dir'] == 'H' else i)
                    grid_usage[(gx, gy)] = grid_usage.get((gx, gy), 0) + 1

            for item in self.placed_words:
                word = item['word_upper']
                # 使用当前循环独特的独立种子，确保算法内部模拟效果与画图时完全对齐一致
                random.seed(f"{word}-{item['x']}-{item['y']}-{session_seed}-{attempt}")
                dynamic_rate = hint_rate + random.uniform(-0.05, 0.05)
                
                letters = []
                for i, char in enumerate(word):
                    if char != '~':
                        gx = item['x'] + (i if item['dir'] == 'H' else 0)
                        gy = item['y'] + (0 if item['dir'] == 'H' else i)
                        
                        # --- 权重分配逻辑 ---
                        weight = 10  # 基础权重 (普通字母)
                        if i == 0: 
                            weight = 60  # 首字母权重
                        else:
                            if grid_usage.get((gx, gy), 0) > 1:
                                weight = 25  # 交叉点权重检查
                            if char in "QXYZẞÄÖÜ":
                                weight = 25  # 稀有字母权重检查 (德语辨识度高的字符)
                        
                        letters.append({'coord': (gx, gy), 'idx': i, 'weight': weight})

                num_to_show = max(1, int(len(letters) * dynamic_rate + 0.5))
                chosen_in_word = []

                # 1. 分段保底：长单词前中后至少各出一个
                if len(letters) >= 6:
                    segments = [
                        [l for l in letters if l['idx'] < len(word)//3],
                        [l for l in letters if len(word)//3 <= l['idx'] < 2*len(word)//3],
                        [l for l in letters if l['idx'] >= 2*len(word)//3]
                    ]
                    for seg in segments:
                        if seg and len(chosen_in_word) < num_to_show:
                            chosen_in_word.append(max(seg, key=lambda x: x['weight']))
                
                # 2. 剩余配额按权重随机填充
                remaining_pool = [l for l in letters if l not in chosen_in_word]
                weighted_pool = []
                for l in remaining_pool: 
                    weighted_pool.extend([l] * l['weight'])
                
                random.shuffle(weighted_pool)
                for candidate in weighted_pool:
                    if len(chosen_in_word) >= num_to_show: break
                    if candidate not in chosen_in_word:
                        # 防连字逻辑
                        if not any(abs(candidate['idx'] - c['idx']) <= 1 for c in chosen_in_word):
                            chosen_in_word.append(candidate)
                        elif len(letters) - len(chosen_in_word) <= 1:
                            chosen_in_word.append(candidate)

                for cand in chosen_in_word:
                    current_hint_coords.add(cand['coord'])

            # ==========================================================
            # 全局扫描：计算跨单词网络交叉重叠后，全局总共出现了多少处“连续3个字及以上显示”
            # ==========================================================
            streak_violations = 0
            for y in range(self.size):
                for x in range(self.size):
                    if self.grid[y][x] == '#' or self.grid[y][x] == '~': 
                        continue
                    if (x, y) in current_hint_coords:
                        # 检查水平方向是否存在连续 3 格显示
                        if (x+1, y) in current_hint_coords and (x+2, y) in current_hint_coords:
                            if self.grid[y][x+1] != '#' and self.grid[y][x+2] != '#':
                                streak_violations += 1
                        # 检查纵向方向是否存在连续 3 格显示
                        if (x, y+1) in current_hint_coords and (x, y+2) in current_hint_coords:
                            if self.grid[y+1][x] != '#' and self.grid[y+2][x] != '#':
                                streak_violations += 1

            # ==========================================================
            # 多目标最高优胜决策筛选：
            # 1. 优先保证收录单词最多 (current_placed_count > max_count)
            # 2. 词数打平时，挑选全局三连字最少的一套方案 (streak_violations < min_streak_violations)
            # ==========================================================
            if (current_placed_count > max_count) or \
               (current_placed_count == max_count and streak_violations < min_streak_violations):
                
                max_count = current_placed_count
                min_streak_violations = streak_violations
                best_state = {
                    'grid': [row[:] for row in self.grid], #深度备份当前的“最佳状态”的棋盘布局
                    'placed': self.placed_words[:],       #保存已放置的单词列表
                    'hints': set(current_hint_coords)      #保存对应的最优提示字格子集合
                }
                
            # 完美收官条件：全词成功塞满 且 全局实现 0 处连字违规，直接提前打破循环
            if max_count == len(raw_data) and min_streak_violations == 0: 
                break

        if best_state: #保存记录到的最佳布局状态
            self.grid = best_state['grid']
            self.placed_words = best_state['placed']
            self.hint_coords = best_state['hints']
            print(f"-> 优选方案质检报告：最大收录单词数: {max_count}，全局 3 连字冲突数降至: {min_streak_violations} 处")
            return max_count
        return 0


def draw_pdf(cw, start, end, hint_rate, all_input_data):
    file_name = f"German_Crossword_Final_R{start}_{end}.pdf"
    c = canvas.Canvas(file_name, pagesize=A4)
    cell_size = 18                  
    offset_x, offset_y = 50, 750    

    for mode in ["quiz", "answer"]:
        c.setFont(FONT_NAME, 14)
        title = "Kreuzworträtsel (Quiz)" if mode=="quiz" else "Lösung (答案页 - 详细对照)"
        c.drawString(50, 810, f"{title} - Nr.{start} bis {end}")

        coord_to_num = {}
        curr_num = 1
        sorted_placed = sorted(cw.placed_words, key=lambda k: (k['y'], k['x']))
        for item in sorted_placed:
            if (item['x'], item['y']) not in coord_to_num:
                coord_to_num[(item['x'], item['y'])] = curr_num
                curr_num += 1
            item['num'] = coord_to_num[(item['x'], item['y'])]

        # 绘制棋盘网格
        for y in range(cw.size):
            for x in range(cw.size):
                char = cw.grid[y][x]
                if char == '#': continue
                
                px, py = offset_x + x * cell_size, offset_y - y * cell_size
                c.setStrokeColorRGB(0.2, 0.2, 0.2)
                c.rect(px, py, cell_size, cell_size, fill=0)

                if char == '~':
                    c.setFillColor(colors.black)
                    c.setFont(FONT_NAME, cell_size * 0.7)
                    c.drawCentredString(px + cell_size/2, py + cell_size * 0.28, "~")
                    continue 

                # 【修改之处】直接读取并对齐在进化算法里锁定的黄金 hint_coords
                # 这样可以 100% 确保 PDF 渲染出的格子绝对是经历过全局 3 连字惩罚优选后的完美形态
                is_hint = (x, y) in cw.hint_coords
                
                if mode == "quiz":
                    if is_hint:
                        c.setFillColorRGB(0.5, 0.5, 0.5)
                        c.setFont(FONT_NAME, cell_size * 0.6)
                        c.drawCentredString(px + cell_size/2, py + cell_size * 0.28, char)
                else:
                    c.setFillColor(colors.grey if is_hint else colors.red)
                    c.setFont(FONT_NAME, cell_size * 0.6)
                    c.drawCentredString(px + cell_size/2, py + cell_size * 0.28, char)

                if (x, y) in coord_to_num:
                    c.setFillColor(colors.black)
                    c.setFont("Helvetica", cell_size * 0.3)
                    c.drawString(px + 1, py + cell_size - 6, str(coord_to_num[(x, y)])) #调整棋盘上首字母位置的序号显示位置

        # 底部列表绘制
        y_ptr = offset_y - (cw.size * cell_size) - 30
        c.setFont(FONT_NAME, 9)
        c.setFillColor(colors.grey) ############################
        if mode == "quiz":
            c.drawString(50, y_ptr, "HINWEISE (线索): W：横向；S：纵向；(i)里面的i表示单词长度"); y_ptr -= 15
            for item in sorted_placed:
                label = f"{item['num']}. [{'W' if item['dir']=='H' else 'S'}] {item['full_info'][1]} ({len(item['word_upper'])})"
                c.drawString(60, y_ptr, label); y_ptr -= 12
                if y_ptr < 50: c.showPage(); y_ptr = 800; c.setFont(FONT_NAME, 9); c.setFillColor(colors.grey)
        else:
            c.drawString(50, y_ptr, "VOKABELN (全5列详细信息 - 红色已入选，黑色孤立):"); y_ptr -= 15
            
            # 绘制答案页已放入游戏的词
            c.setFillColor(colors.red)
            for itr_lst1 in sorted_placed:
                info_text = " | ".join([str(itr_lst2) for itr_lst2 in itr_lst1['full_info']])
                info_text = re.sub(r'[| ]+$', '', info_text)
                c.drawString(60, y_ptr, f"{itr_lst1['num']}. {info_text}")
                y_ptr -= 12
                if y_ptr < 50: c.showPage(); y_ptr = 800; c.setFont(FONT_NAME, 9); c.setFillColor(colors.red)

            # 绘制孤立词
            c.setFillColor(colors.black)    
            for wu, fd in all_input_data:
                if wu not in [itr_lst['word_upper'] for itr_lst in sorted_placed]:
                    info_text = " | ".join([str(i) for i in fd])
                    info_text = re.sub(r'[| ]+$', '', info_text)
                    c.drawString(60, y_ptr, f"○ {info_text} ")
                    y_ptr -= 12
                    if y_ptr < 50: c.showPage(); y_ptr = 800; c.setFont(FONT_NAME, 9); c.setFillColor(colors.black)

        c.showPage()
    c.save()


if __name__ == "__main__":
    try:
        df = pd.read_excel(GermansWordFile, header=None)         #首行为自定义表头，读取数据从第2行开始
        max_idx = len(df) - 1
        subset = df.iloc[min(START_ROW-1, max_idx) : min(END_ROW-1, max_idx) + 1]   #得到excel记录转换来的数据帧
        
        all_input_data = []                         #存储待输入棋盘的大写单词
        for _, row in subset.iterrows():            #数据帧逐行迭代访问
            if pd.notna(row[0]):                    #每个单词列(首列)不为空
                full_info = []                      #存储每个单词及其后续的注释(对应excel的1~5列)
                for i in range(5):
                    # 修改：即使是提取数据阶段也应用 super_clean_text
                    cleaned = super_clean_text(row[i]) if i < len(row) else ""
                    full_info.append(cleaned)
                
                wu = clean_for_grid(str(row[0]))
                if len(wu) > 1:
                    strMaxLlimit = 20                  #限制单词长度不能超过20个字母，否则截短
                    wu = wu[:strMaxLlimit] if len(wu)>20 else wu
                    all_input_data.append((wu, full_info))
        
        if all_input_data:
            engine = CrosswordEngine(size=30)
            # 【修改之处】生成全局统一固定的唯一种子，并将 HINT_RATE 和种子透传给 engine 内部做质检
            session_seed = random.randint(0, 10**6) 
            success = engine.generate_best_layout(
                all_input_data, 
                iterations=ITERATIONS, 
                hint_rate=HINT_RATE, 
                session_seed=session_seed
            )
            print(f"完成！放入 {success}/{len(all_input_data)} 个词。")
            draw_pdf(engine, START_ROW, END_ROW, HINT_RATE, all_input_data)
        else:
            print("选定范围内无有效数据。")
    except Exception as e:
        import traceback
        print(f"运行失败: {e}")
        traceback.print_exc()