# DOCX 合同审查操作 SOP——失败经验与技术规范

> 本文件记录在原文件上叠加审查修改的完整操作流程，包含历次失败经验、技术陷阱与应做/不应做清单。适用于 contract-copilot skill 自动化审查及人工辅助审查场景。

---

## 一、核心原则：原文件叠加修改

### 应做

1. **直接打开原始 DOCX 文件**进行修改，使用 `python-docx` 的 `Document(src_path)` 打开原文件。
2. **所有修改通过 Word 原生修订功能完成**：插入用 `<w:ins>`，删除用 `<w:del>`，批注用 `<w:comment>` + `<w:commentRangeStart/End>` + `<w:commentReference>`。
3. **继承原文档全部格式**：新插入文本必须设置与原文一致的字体名称（如 `仿宋_GB2312`）和字体大小（如 `w:sz=28`，即 14pt 四号字），通过 `copy.deepcopy(rPr)` 或手动构造完整 `rPr` 实现。
4. **新增条款精确定位插入**：使用 `body.insert(pos, new_p)` 在目标位置（如"以下无正文"段落之前）插入，禁止使用 `body.append()` 追加到文档末尾。
5. **run 级精确替换**：找到包含目标文本的 run，拆分为 `before + [del old + ins new] + after` 三个元素组，避免破坏句子结构和上下文。
6. **保留全部已有批注和修订**：不得修改、删除、接受或拒绝他人已有的批注与修订痕迹。仅以当前审查人身份新增内容。
7. **审查人身份隔离**：所有新增的 `w:ins`、`w:del`、`w:comment` 必须使用当前审查人署名（如 `Vang`），与已有审查人（如 `理论科`、`星星上的花`）严格区分。

### 不应做

1. **禁止从零创建新文档**：不得将原文文本复制到新建的 Document 中再修改，这会丢失全部原始格式（字体、段落、编号、缩进、行距、下划线）。
2. **禁止使用 `body.append()` 插入新条款**：这会将新增内容追加到文档最末尾（签署栏之后），而非逻辑正确的位置。
3. **禁止只设置字体名称不设置字体大小**：仅设置 `w:rFonts` 而不设 `w:sz` 和 `w:szCs`，会导致新插入文本字体大小与原文不一致（如原文 14pt，新文本显示为默认 10pt 或 12pt）。
4. **禁止修改原有 comments.xml 内容**：原有批注的 `w:id`、`w:author`、`w:date`、文本内容均不得改动。
5. **禁止接受或拒绝原有修订**：不得对已有 `w:ins`/`w:del` 执行 accept/reject 操作。
6. **禁止在 `OxmlElement` 上调用 `set(qn('xmlns:w'), ...)` 设置命名空间**：这会导致 python-docx 内部 nsmap 查找失败，抛出 `KeyError: 'xmlns'`。
7. **禁止使用 `OxmlElement(f'{{{uri}}}')` 语法传入完整命名空间 URI**：应使用 `qn('w:xxx')` 简写形式，或用 `etree.SubElement(parent, f'{{{uri}}}tag')`。

---

## 二、.doc 文件转换

### 应做

1. **检测文件是否为真 docx**：用 `open(path, 'rb').read(4)` 检查文件头是否为 `PK\x03\x04`（ZIP 格式）。
2. **指导用户用 WPS Office "另存为"**：`.doc` 文件必须通过 WPS 或 Word 的"文件→另存为→Word 文档(.docx)"功能转换。
3. **转换后二次验证**：检查转换后的文件头是否为 ZIP 格式，确认 `word/document.xml`、`word/comments.xml` 等内部结构完整。

### 不应做

1. **禁止仅修改文件扩展名**：将 `.doc` 改为 `.docx` 不会转换内部格式，文件仍是 OLE/复合文档（Composite Document File V2），python-docx 无法读取。
2. **禁止使用 macOS `textutil` 转换**：`textutil -convert docx` 会丢失全部批注和修订痕迹，且生成的文件结构与标准 docx 不同。
3. **禁止使用 LibreCLI 等命令行工具转换**：同样可能丢失批注和修订。

---

## 三、python-docx save() 的批注覆盖陷阱

### 问题描述

`python-docx` 的 `doc.save()` 方法会重新生成 `word/comments.xml` part，导致原有批注（其他审查人添加的）丢失。

### 解决方案

1. **先调用 `doc.save(output)` 保存文档**。
2. **从原始源 ZIP 中读取 `word/comments.xml`**。
3. **合并批注**：将原有批注（`w:id` 不在新生成 comments.xml 中的）追加到新 comments_root。
4. **用 `zipfile` 重写 ZIP**：读取 output 文件全部内容，替换 `word/comments.xml`，重新写入。

### 代码模板

```python
# 1. python-docx 操作
doc = Document(src)
# ... 添加批注和修订 ...
doc.save(output)

# 2. 合并原有批注
orig_z = zipfile.ZipFile(src)
orig_comments_root = etree.fromstring(orig_z.read('word/comments.xml'))
orig_z.close()

existing_ids = {c.get(qn('w:id')) for c in comments_root.findall(f'{{{W}}}comment')}
for orig_c in orig_comments_root.findall(f'{{{W}}}comment'):
    if orig_c.get(qn('w:id')) not in existing_ids:
        comments_root.append(copy.deepcopy(orig_c))

new_comments_xml = etree.tostring(comments_root, xml_declaration=True,
                                   encoding='UTF-8', standalone=True)

# 3. 替换 ZIP 内的 comments.xml
with zipfile.ZipFile(output, 'r') as zin:
    names = zin.namelist()
    data = {name: zin.read(name) for name in names}
data['word/comments.xml'] = new_comments_xml
with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name in names:
        zout.writestr(name, data[name])
```

---

## 四、完整 rPr 构造规范

新插入文本的 `w:rPr` 必须包含以下全部属性，缺一不可：

| 属性 | 说明 | 示例值 |
|------|------|--------|
| `w:rFonts/w:eastAsia` | 中文字体 | `仿宋_GB2312` |
| `w:rFonts/w:ascii` | 西文字体 | `仿宋_GB2312` |
| `w:rFonts/w:hAnsi` | 西文字体（高版本） | `仿宋_GB2312` |
| `w:sz/w:val` | 字号（半磅） | `28`（=14pt=四号） |
| `w:szCs/w:val` | 字号（复杂脚本） | `28` |

### 构造函数

```python
def make_rPr(font_name='仿宋_GB2312', bold=False):
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rPr.append(rFonts)
    if bold:
        rPr.append(OxmlElement('w:b'))
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '28')
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), '28')
    rPr.append(szCs)
    return rPr
```

---

## 五、新增条款位置定位

### 定位锚点

新增条款（如"通知与送达"、"合同解除"）应插入在以下锚点之前：

- `"以下无正文"` 段落
- 签署栏（"甲方：（盖章）"）段落
- 文档最后一个段落

### 定位代码

```python
# 找到"以下无正文"段落
no_text_para = None
for p in doc.paragraphs:
    if '以下无正文' in p.text:
        no_text_para = p
        break

# 在该段落前插入新段落
body = doc.element.body
no_text_pos = list(body).index(no_text_para._p)
for idx, (text, bold) in enumerate(new_sections):
    new_p = build_paragraph_with_ins(text, bold)
    body.insert(no_text_pos + idx, new_p)
```

### 常见错误

- 使用 `doc.add_paragraph()` → 追加到文档最末尾
- 使用 `body.append(new_p)` → 同上
- 在 `insert_after` 之后用 `doc.paragraphs` 索引定位 → 索引已因插入而偏移，应直接用 XML 元素引用

---

## 六、people.xml 更新

新审查人加入时，须更新 `word/people.xml`：

```python
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
people_root = etree.fromstring(data['word/people.xml'])
# 检查是否已存在
exists = any(p.get(f'{{{W15}}}author') == AUTHOR
              for p in people_root.findall(f'{{{W15}}}person'))
if not exists:
    person = etree.SubElement(people_root, f'{{{W15}}}person')
    person.set(f'{{{W15}}}author', AUTHOR)
    presence = etree.SubElement(person, f'{{{W15}}}presenceInfo')
    presence.set(f'{{{W15}}}providerId', 'None')
    presence.set(f'{{{W15}}}userId', AUTHOR)
```

---

## 七、验证清单

生成审查版后，必须执行以下验证：

### 7.1 批注验证
- [ ] 原有批注数量不变（对比源文件 comments.xml 的 `w:comment` 数量）
- [ ] 原有批注作者、日期、文本未被修改
- [ ] 新增批注作者为当前审查人（如 `Vang`）
- [ ] 新增批注 id 不与原有 id 冲突

### 7.2 修订验证
- [ ] 原有 `w:ins` 数量不变，作者不变
- [ ] 原有 `w:del` 数量不变，作者不变
- [ ] 新增 `w:ins`/`w:del` 作者为当前审查人

### 7.3 字体验证
- [ ] 所有新增 `w:ins` 内 `w:r` 的 `rPr` 包含 `w:sz`（字号）
- [ ] 所有新增 `w:del` 内 `w:r` 的 `rPr` 包含 `w:sz`
- [ ] 字体名称与原文一致（如 `仿宋_GB2312`）

### 7.4 位置验证
- [ ] 新增条款位于"以下无正文"之前
- [ ] 新增条款不在签署栏之后
- [ ] 原有条款顺序未被打乱

### 7.5 people.xml 验证
- [ ] 原有审查人保留
- [ ] 当前审查人已添加

### 验证注意事项

- `doc.paragraphs[i].text` **不读取** `w:ins` 元素内的文本，验证含修订段落内容时须用 `p._p.iter('{W}t')` 遍历所有 `w:t` 元素。
- `p.runs` 只返回段落直接子级的 `w:r`，不包含 `w:ins`/`w:del` 内的 `w:r`，验证修订内字体时须用 `p._p.findall('.//{W}ins')` 递归查找。

---

## 八、历次失败经验汇总

| 序号 | 失败现象 | 根因 | 修复方案 |
|------|---------|------|---------|
| 1 | 格式全部丢失（字体、段落、编号消失） | 从零创建新文档而非在原文件上修改 | 直接用 `Document(src)` 打开原文件操作 |
| 2 | 原有批注丢失 | python-docx `save()` 重新生成 comments part | save 后从原始 ZIP 读取 comments.xml 合并 |
| 3 | 修订插入位置破坏句子结构 | del+ins 插在原 run 前面，未拆分 run | run 级精确拆分：before + [del + ins] + after |
| 4 | 批注作者显示为"合同审查AI" | 硬编码了泛化名称 | 统一使用 `Vang` |
| 5 | 条款内容过于简略 | 未按六要素闭环标准撰写 | 按"通知→核实→处理→费用→时限→违约后果"六要素完善 |
| 6 | 新增条款跑到文档最后 | 使用 `body.append()` 而非 `body.insert(pos)` | 精确定位"以下无正文"段落，用 `body.insert()` 插入 |
| 7 | 新增文本字体大小与原文不一致 | rPr 只设了 `w:rFonts` 没设 `w:sz`/`w:szCs` | `make_rPr()` 函数统一设置字体名+大小 |
| 8 | `KeyError: 'xmlns'` | 在 OxmlElement 上调用 `set(qn('xmlns:w'), ...)` | 删除命名空间 set 操作 |
| 9 | `UnboundLocalError: 'zipfile'` | 函数内部 `import zipfile` 与模块级冲突 | 删除函数内部导入，使用顶层导入 |
| 10 | .doc 改扩展名为 .docx 后读取失败 | 内部仍为 OLE 格式，非 ZIP | 指导用户用 WPS"另存为"转换 |
| 11 | 验证脚本报段落文本为空 | `paragraph.text` 不读取 `w:ins` 内文本 | 用 `p._p.iter('{W}t')` 遍历提取 |
| 12 | `etree.SubElement` 使用 `{{{uri}}}` 语法报错 | OxmlElement 不支持完整 URI 语法 | 用 `qn('w:xxx')` 或 `etree.SubElement(parent, f'{{{uri}}}tag')` |
