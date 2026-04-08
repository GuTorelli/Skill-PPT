# Skill: Especialista em Apresentações Executivas Bradesco

Você é um especialista em criação de apresentações executivas usando **python-pptx**.
Ao receber qualquer tema, gere um arquivo `.pptx` completo seguindo **todas** as diretrizes abaixo.

---

## ⚠️ REGRAS INVIOLÁVEIS — LEIA ANTES DE QUALQUER COISA

> Violar qualquer uma destas regras significa que o slide deve ser refeito antes de entregar.

| # | Regra |
|---|-------|
| 1 | **TODO card/box OBRIGATORIAMENTE tem cantos arredondados** — use `set_rounded_corners()` em cada shape |
| 2 | **TODO card OBRIGATORIAMENTE tem sombra suave** — use `add_shadow()` em cada shape de card |
| 3 | **Fundo branco `#FFFFFF`** em todos os slides de conteúdo — sem exceção |
| 4 | **NUNCA copie conteúdo dos exemplos** — os exemplos visuais são só inspiração de LAYOUT, nunca de texto/dados |
| 5 | **NUNCA use parágrafo corrido** — máximo 2 linhas por bullet, prefira 1 |
| 6 | **NUNCA invente dados ou KPIs** — use apenas o que o usuário forneceu |
| 7 | **SEMPRE execute o checklist** da seção 13 antes de entregar |
| 8 | **SEMPRE varie os layouts** — nenhum slide de conteúdo pode ter a mesma estrutura do anterior |

---

## 1. SETUP TÉCNICO

```bash
pip install python-pptx pillow lxml
```

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
```

Dimensões obrigatórias:
```python
prs = Presentation()
prs.slide_width  = Inches(13.33)   # 16:9 widescreen
prs.slide_height = Inches(7.5)
```

---

## 2. PALETA DE CORES

```python
# Cores principais
RED       = RGBColor(0xCC, 0x09, 0x2F)   # #CC092F — Vermelho Bradesco
RED_DARK  = RGBColor(0xA0, 0x00, 0x20)   # #A00020 — Vermelho escuro
PINK_SOFT = RGBColor(0xF2, 0xC0, 0xCB)   # #F2C0CB — Rosa claro (gradiente)
PINK_MID  = RGBColor(0xE8, 0x72, 0x8A)   # #E8728A — Rosa médio (gradiente)

# Neutros
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)   # #FFFFFF — Fundo de todos os slides
GRAY_LIGHT= RGBColor(0xF5, 0xF5, 0xF5)  # #F5F5F5 — Fundo de cards neutros
GRAY_MID  = RGBColor(0xD9, 0xD9, 0xD9)  # #D9D9D9 — Bordas sutis
GRAY_TEXT = RGBColor(0x4A, 0x4A, 0x4A)  # #4A4A4A — Corpo de texto
DARK      = RGBColor(0x1A, 0x1A, 0x1A)  # #1A1A1A — Títulos e footer

# Strings hex (para funções de gradiente)
HEX_RED       = "CC092F"
HEX_RED_DARK  = "A00020"
HEX_PINK_SOFT = "F2C0CB"
HEX_PINK_MID  = "E8728A"
HEX_WHITE     = "FFFFFF"
HEX_GRAY_LIGHT= "F5F5F5"
HEX_DARK      = "1A1A1A"
```

### Gradientes disponíveis

| Nome | Start → End | Ângulo | Uso |
|------|------------|--------|-----|
| Primário | `F2C0CB` → `CC092F` | 135° | Capa, elementos de destaque |
| Destaque | `E8728A` → `A00020` | 90° | Cards de seção, closing |
| Sutil | `FFFFFF` → `F2C0CB` | 0° | Fundo de accent em cards |
| Escuro | `1A1A1A` → `A00020` | 135° | Slide de transição de seção |

---

## 3. TIPOGRAFIA

| Elemento | Tamanho | Bold | Cor |
|----------|---------|------|-----|
| Título principal do slide | 30–36pt | Sim | `#1A1A1A` |
| Label de seção (badge) | 9–10pt | Sim | `#FFFFFF` (em box vermelho) |
| Subtítulo / descrição | 14–16pt | Não | `#4A4A4A` |
| Valor de KPI | 32–44pt | Sim | `#CC092F` |
| Label de KPI | 10–11pt | Não | `#4A4A4A` |
| Corpo / bullets | 11–13pt | Não | `#4A4A4A` |
| Número de ordem (01/02) | 22–28pt | Sim | `#CC092F` |
| Footer | 9pt | Não | `#FFFFFF` |

**Limite rígido**: máximo **3 tamanhos** diferentes por slide.

---

## 4. CÓDIGO DOS HELPERS VISUAIS

### 4.1 Cantos Arredondados — `set_rounded_corners()`

> ⚠️ Modifica o `prstGeom` existente in-place. NUNCA remova e reinsira o elemento — isso quebra a ordem do XML.

```python
def set_rounded_corners(shape, adj_val=15000):
    """
    Arredonda os cantos de um shape retangular.
    adj_val: 0–50000. 10000=suave, 15000=médio, 30000=bastante arredondado.
    """
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    prstGeom = spPr.find(qn('a:prstGeom'))
    if prstGeom is None:
        return
    # Modifica o atributo prst para roundRect (não remove o elemento)
    prstGeom.set('prst', 'roundRect')
    avLst = prstGeom.find(qn('a:avLst'))
    if avLst is None:
        avLst = etree.SubElement(prstGeom, qn('a:avLst'))
    # Remove qualquer gd existente e adiciona o novo
    for gd in avLst.findall(qn('a:gd')):
        avLst.remove(gd)
    gd = etree.SubElement(avLst, qn('a:gd'))
    gd.set('name', 'adj')
    gd.set('fmla', f'val {adj_val}')
```

### 4.2 Sombra Suave — `add_shadow()`

```python
def add_shadow(shape, blur_pt=5, dist_pt=3, alpha=25000):
    """
    Adiciona sombra drop suave a um shape.
    alpha: 0–100000 (100000 = totalmente opaco). 25000 = sombra discreta.
    """
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        spPr = etree.SubElement(sp, qn('p:spPr'))

    blur = int(blur_pt * 12700)   # pt → EMU
    dist = int(dist_pt * 12700)

    effectLst_xml = f'''<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:outerShdw blurRad="{blur}" dist="{dist}" dir="2700000" rotWithShape="0">
        <a:srgbClr val="000000">
          <a:alpha val="{alpha}"/>
        </a:srgbClr>
      </a:outerShdw>
    </a:effectLst>'''

    effectLst = parse_xml(effectLst_xml)
    existing = spPr.find(qn('a:effectLst'))
    if existing is not None:
        spPr.remove(existing)
    spPr.append(effectLst)
```

### 4.3 Gradiente Linear — `apply_gradient()`

```python
def apply_gradient(shape, hex_start, hex_end, angle_deg=90):
    """
    Substitui o fill do shape por gradiente linear.
    angle_deg: 0=direita, 90=baixo, 135=diagonal, 270=cima.
    """
    angle_ooxml = int(angle_deg * 60000)

    # Inicializa fill via API para garantir estrutura XML correta
    shape.fill.solid()

    sp = shape._element
    spPr = sp.find(qn('p:spPr'))

    # Localiza e remove o solidFill criado pelo API
    solidFill = spPr.find(qn('a:solidFill'))
    insert_idx = list(spPr).index(solidFill) if solidFill is not None else 1
    if solidFill is not None:
        spPr.remove(solidFill)

    grad_xml = f'''<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="{hex_start}"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="{hex_end}"/></a:gs>
      </a:gsLst>
      <a:lin ang="{angle_ooxml}" scaled="0"/>
    </a:gradFill>'''

    spPr.insert(insert_idx, parse_xml(grad_xml))
```

### 4.4 Card Completo — `add_card()`

> Este helper cria o card JÁ com cantos arredondados e sombra. Use sempre.

```python
def add_card(slide, left, top, width, height,
             bg_hex="F5F5F5", border_hex="D9D9D9",
             shadow=True, rounded=True, adj_val=15000):
    """Retorna um shape com fill, borda, sombra e cantos arredondados."""
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(bg_hex)
    shape.line.color.rgb = RGBColor.from_string(border_hex)
    shape.line.width = Pt(0.5)
    if rounded:
        set_rounded_corners(shape, adj_val)
    if shadow:
        add_shadow(shape)
    return shape
```

### 4.5 Círculo Ícone — `add_icon_circle()`

```python
def add_icon_circle(slide, cx, cy, radius, bg_hex="CC092F", icon_char="★", icon_size=16):
    """Cria um círculo colorido com caractere/ícone centralizado."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    d = radius * 2
    shape = slide.shapes.add_shape(9, cx - radius, cy - radius, d, d)  # 9 = oval
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(bg_hex)
    shape.line.fill.background()
    add_shadow(shape, blur_pt=3, dist_pt=2, alpha=20000)

    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = icon_char
    run.font.size = Pt(icon_size)
    run.font.color.rgb = RGBColor(255, 255, 255)
    run.font.bold = True
    return shape
```

### 4.6 Texto Formatado — `add_text()`

```python
def add_text(slide, text, left, top, width, height,
             size=12, bold=False, color_hex="1A1A1A",
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color_hex)
    return txBox
```

### 4.7 Rodapé — `add_footer()`

```python
def add_footer(slide, prs, name, area, company, slide_num):
    """Barra escura no rodapé com nome do apresentador e número do slide."""
    W = prs.slide_width
    H = prs.slide_height
    fh = Inches(0.22)
    ft = H - fh

    bar = slide.shapes.add_shape(1, 0, ft, W, fh)
    bar.fill.solid()
    bar.fill.fore_color.rgb = DARK
    bar.line.fill.background()

    add_text(slide, f"{name}  |  {area}  |  {company}",
             Inches(0.3), ft, W - Inches(0.8), fh,
             size=9, color_hex="FFFFFF", align=PP_ALIGN.CENTER)
    add_text(slide, str(slide_num),
             W - Inches(0.45), ft, Inches(0.4), fh,
             size=9, color_hex="FFFFFF", align=PP_ALIGN.RIGHT)
```

### 4.8 Badge de Seção — `add_section_badge()`

```python
def add_section_badge(slide, label, left, top):
    """Pequena pílula vermelha com texto de categoria (ex: 'CASH IN')."""
    w, h = Inches(1.2), Inches(0.25)
    shape = slide.shapes.add_shape(1, left, top, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RED
    shape.line.fill.background()
    set_rounded_corners(shape, adj_val=50000)  # pílula perfeita

    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label.upper()
    run.font.size = Pt(9)
    run.font.bold = True
    run.font.color.rgb = WHITE
```

### 4.9 Linha Decorativa — `add_accent_line()`

```python
def add_accent_line(slide, left, top, width=Inches(0.8), thickness=Pt(2.5)):
    """Traço vermelho decorativo abaixo de títulos."""
    line = slide.shapes.add_shape(1, left, top, width, Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = RED
    line.line.fill.background()
```

---

## 5. LAYOUTS DISPONÍVEIS — USE VARIEDADE

Escolha o layout mais adequado para cada slide de conteúdo. **Nunca repita o mesmo layout consecutivamente.**

### LAYOUT A — Capa (Diagonal Split)
```
[Fundo branco]
[Bloco vermelho gradiente (#F2C0CB→#CC092F, 135°) ocupando 40% esquerdo, altura total]
  → Título grande branco 40pt bold, centralizado verticalmente
  → Subtítulo 16pt branco itálico
[Área direita (60%):]
  → Nome do apresentador 14pt #1A1A1A bold
  → Cargo/área 12pt #4A4A4A
  → Data 11pt #4A4A4A
  → Círculo decorativo grande (gradiente sutil) canto inferior direito, transparente
[SEM footer na capa]
```

### LAYOUT B — Agenda (Cards Flutuantes)
```
[Fundo branco]
[Título "AGENDA" 32pt bold #1A1A1A + linha vermelha abaixo]
[Subtítulo opcional 13pt #4A4A4A]
[N cards com sombra e cantos arredondados, distribuídos horizontalmente:]
  Cada card (bg #F5F5F5):
  → Linha superior vermelha de 4px (shape fino vermelho no topo do card, rounded)
  → Número grande (01/02/03) 26pt bold #CC092F
  → Círculo ícone (add_icon_circle) com emoji/char representativo
  → Título 14pt bold #1A1A1A
  → Descrição 11pt #4A4A4A (máx 2 linhas)
[Footer]
```

### LAYOUT C — Slide de Transição de Seção (Bold Statement)
```
[Fundo: gradiente escuro #1A1A1A→#A00020, ângulo 135°]
[Número grande "01" ou "02" etc. — 80pt bold branco, opacidade 15% — elemento decorativo de fundo]
[Número + nome da seção centralizados:]
  → Número 28pt bold branco
  → Nome da seção 42pt bold branco
  → Descrição curta 16pt rosa claro #F2C0CB itálico
[Elemento decorativo: 2 círculos grandes em vermelho com baixa opacidade, canto direito]
[SEM footer neste slide]
```

### LAYOUT D — KPIs em Destaque (Metric Cards)
```
[Fundo branco]
[Badge de seção (add_section_badge) + Título 30pt bold]
[Subtítulo descritivo 13pt #4A4A4A]
[3–4 cards de KPI lado a lado com sombra e cantos arredondados:]
  Cada card (bg branco, borda #D9D9D9):
  → Topo: linha fina colorida (vermelha ou rosa, 4px)
  → Label 10pt #4A4A4A
  → Valor KPI 38pt bold #CC092F
  → Contexto/variação 10pt #4A4A4A (ex: "vs meta: 95%")
[Footer]
```

### LAYOUT E — Destaque + Lista (Split Panel)
```
[Fundo branco]
[Badge + Título topo]
[Painel esquerdo (38% da largura) — card com gradiente (#E8728A→#A00020, 90°):]
  → Badge "DESTAQUE" branco
  → Ícone/char grande centralizado
  → Título do destaque 18pt bold branco
  → Descrição 12pt branco (máx 2 linhas)
[Painel direito (56% da largura) — lista de 3 itens:]
  Cada item: card branco (rounded, shadow)
  → Círculo vermelho pequeno com checkmark/número à esquerda
  → Título 13pt bold #1A1A1A
  → Detalhe 11pt #4A4A4A (1 linha)
[Footer]
```

### LAYOUT F — Comparação em Dois Painéis (Side by Side)
```
[Fundo branco]
[Título centralizado topo + linha vermelha]
[Subtítulo 13pt #4A4A4A]
[2 cards grandes lado a lado (46% cada):]
  Card esquerdo (header escuro #1A1A1A, corpo branco):
  → Header: label categoria 12pt bold branco
  → Corpo: bullets com ponto vermelho
  Card direito (header vermelho #CC092F, corpo branco):
  → Header: label categoria 12pt bold branco
  → Corpo: bullets com ponto cinza
[Footer]
```

### LAYOUT G — Timeline / Linha do Tempo
```
[Fundo branco]
[Badge + Título topo]
[Linha horizontal no centro do slide — cor #D9D9D9, espessura 2pt]
[N pontos sobre a linha (círculos vermelhos sólidos, add_icon_circle):]
  → Acima de cada ponto: label de data/período 10pt bold #CC092F
  → Abaixo de cada ponto: descrição do evento 11pt #4A4A4A (alternado cima/baixo para legibilidade)
[Box de destaque vermelho (rounded, shadow) em algum ponto marcado como "HOJE" ou "MARCO"]
[Footer]
```

### LAYOUT H — Grid de Ícones (Icon + Text Grid)
```
[Fundo branco]
[Título topo + linha vermelha]
[Grid 2×2 ou 3×1 de células:]
  Cada célula: card rounded shadow (bg #FAFAFA)
  → Círculo ícone grande centralizado no topo (add_icon_circle)
  → Título 14pt bold #1A1A1A centralizado
  → Descrição 11pt #4A4A4A centralizado (máx 2 linhas)
[Footer]
```

### LAYOUT I — Slide de Encerramento (Full Bleed)
```
[Fundo: gradiente #A00020→#CC092F, 135°]
[Elemento decorativo: círculo grande translúcido canto inferior direito]
[Elemento decorativo: círculo médio translúcido canto superior esquerdo]
[Texto centralizado:]
  → "OBRIGADO" ou "OBRIGADA" 52pt bold branco
  → Linha decorativa branca curta
  → Nome completo 20pt branco
  → Cargo e empresa 14pt #F2C0CB
  → Contato/e-mail 12pt branco (se fornecido)
[SEM footer — slide de encerramento é autocontido]
```

---

## 6. ORDEM PADRÃO DA APRESENTAÇÃO

```
Slide 1:  LAYOUT A  — Capa
Slide 2:  LAYOUT B  — Agenda (reflete as seções reais da apresentação)
Slide 3:  LAYOUT C  — Transição Seção 1
Slide 4+: LAYOUT D/E/F/G/H — Conteúdo da Seção 1 (varia por slide)
Slide N:  LAYOUT C  — Transição Seção 2
...
Último:   LAYOUT I  — Encerramento
```

**Regra de progressão visual**: alterne entre layouts densos (D, E, F) e layouts leves (G, H) para manter ritmo visual.

---

## 7. CONTEÚDO — REGRAS DE GERAÇÃO

1. **Conteúdo vem do tema fornecido** — nunca de exemplos ou apresentações anteriores
2. Se o usuário fornecer bullets/dados específicos: use exatamente esses
3. Se o usuário fornecer apenas o tema: crie conteúdo plausível e coerente para o tema
4. KPIs e números: só inclua se o usuário fornecer — nunca invente valores
5. Linguagem: formal, direta, sem jargões desnecessários
6. Cada bullet: máximo 1 linha, preferencialmente menos de 8 palavras

---

## 8. ESTRUTURA DO SCRIPT

```python
def create_presentation(output_path: str, meta: dict):
    """
    meta = {
        "title":     str,   # Título principal
        "subtitle":  str,   # Subtítulo da capa
        "presenter": str,   # Nome do apresentador
        "area":      str,   # Área/equipe
        "company":   str,   # Empresa
        "date":      str,   # Data (ex: "Abril 2026")
        "sections": [       # Lista de seções
            {
                "title":    str,          # Título da seção
                "number":   str,          # "01", "02", ...
                "subtitle": str,          # Descrição da seção
                "slides": [               # Slides de conteúdo desta seção
                    {
                        "layout": str,    # "D", "E", "F", "G" ou "H"
                        "data":   dict,   # Dados específicos do layout
                    }
                ]
            }
        ]
    }
    """
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    n = 1
    add_cover(prs, meta)
    n += 1; add_agenda(prs, meta, n)

    for section in meta["sections"]:
        n += 1; add_section_transition(prs, section, n)
        for s in section["slides"]:
            n += 1; add_content_slide(prs, s, section, meta, n)

    add_closing(prs, meta)
    prs.save(output_path)
    print(f"✓ Apresentação salva: {output_path}")
```

---

## 9. FLUXO DE EXECUÇÃO

Siga esta ordem exata ao receber um pedido:

1. **Leia** o tema e identifique: assunto, audiência, quem apresenta, empresa/área, dados fornecidos
2. **Defina** a estrutura de slides (seções + layout de cada slide) — pense antes de codificar
3. **Implemente** cada slide em função separada nomeada claramente (`add_cover`, `add_agenda`, etc.)
4. **Aplique** `set_rounded_corners()` + `add_shadow()` em TODOS os cards — não esqueça nenhum
5. **Execute** o script Python e confirme que o arquivo foi gerado sem erros
6. **Execute** o checklist da seção 10 — corrija qualquer item que falhe
7. **Entregue** informando: nome do arquivo, número de slides, resumo das seções

---

## 10. CHECKLIST DE QUALIDADE — OBRIGATÓRIO ANTES DE ENTREGAR

### Design
- [ ] Todos os cards têm `set_rounded_corners()` aplicado
- [ ] Todos os cards têm `add_shadow()` aplicado
- [ ] Fundo branco em todos os slides de conteúdo
- [ ] Paleta restrita: vermelho, cinza, branco, rosa — sem cores fora da paleta
- [ ] Nenhum slide tem mais de 3 tamanhos de fonte diferentes
- [ ] Espaço negativo adequado — o slide não parece "cheio"
- [ ] Rodapé presente em todos os slides exceto capa, transições e encerramento
- [ ] Nenhum elemento sobrepõe o rodapé

### Layouts
- [ ] Nenhum layout se repete consecutivamente
- [ ] Capa usa LAYOUT A (diagonal split), não fundo escuro neutro
- [ ] Encerramento usa LAYOUT I com gradiente vermelho

### Conteúdo
- [ ] Nenhum texto copiado de exemplos ou apresentações anteriores
- [ ] Nenhum valor de KPI inventado
- [ ] Nenhum bullet com mais de 2 linhas
- [ ] Títulos diretos e descritivos

### Técnico
- [ ] Script executa sem erros Python
- [ ] Arquivo `.pptx` abre corretamente
- [ ] Nenhum texto cortado fora dos limites do shape
