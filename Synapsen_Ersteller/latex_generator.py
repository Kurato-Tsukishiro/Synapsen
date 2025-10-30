from pathlib import Path
import PDFMargeHelper as Helper


def create_latex_source(notes_info, config, title, paper_size="A4"):
    """
    ノート情報と設定から、統合PDFの元となるLaTeXソースコードを生成する。
    """
    def tex_escape(text):
        # LaTexで特殊文字として扱われる文字をエスケープする
        return str(text).replace('&', r'\&').replace('%', r'\%').replace('$', r'\$') \
                        .replace('#', r'\#').replace('_', ' ') \
                        .replace('{', r'\{').replace('}', r'\}').replace('~', r'\textasciitilde{}') \
                        .replace('^', r'\textasciicircum{}').replace('\\', r'\textbackslash{}') \
                        .replace('<', r'\textless{}').replace('>', r'\textgreater{}') \
                        .replace('|', r'\textbar{}')

    # 設定をconfig辞書から取得
    latex_font = config.get('latex_font', 'Yu Gothic')
    latex_author = config.get('latex_author', 'Your Name')
    key_icons = config.get('key_icons', {})
    key_colors = config.get('key_colors', {})

    # paper_size に応じてdocumentclassのオプションを変更
    paper_option =\
        "a5paper" if paper_size.upper() == "A5" else "a4paper"
    font_size =\
        "10pt" if paper_size.upper() == "A5" else "11pt"
    margin_setting =\
        "margin=2cm" if paper_size.upper() == "A5" else "margin=2.5cm"

    preamble = fr"""
\documentclass[{paper_option}, {font_size}, lualatex]{{ltjsarticle}}
\usepackage{{luatexja-fontspec}}
\usepackage{{fancyhdr}}
\usepackage{{imakeidx}}
\usepackage{{hyperref}}
\usepackage{{pdfpages}}
\usepackage{{multido}}
\usepackage{{xcolor}}
\usepackage[
    {paper_option},
    {margin_setting},
    headheight=2.5cm,
    headsep=1cm,
    footskip=1.5cm
]{{geometry}}

\setmainjfont{{{latex_font}}} \setsansjfont{{{latex_font}}}
\pagestyle{{fancy}} \fancyhf{{}} \renewcommand{{\headrulewidth}}{{0.4pt}}
\renewcommand{{\footrulewidth}}{{0.4pt}} \fancyfoot[C]{{\thepage}}
\hypersetup{{
    colorlinks=true, linkcolor=blue, urlcolor=blue, pdftitle={{{tex_escape(title)}}},
    pdfauthor={{{tex_escape(latex_author)}}}, bookmarks=true, bookmarksnumbered=true, bookmarkstype=toc
}}

\makeindex[name=tags, title=タグ索引, intoc, columns=1]
\makeindex[name=cpkeys, title=Index Key 索引, intoc, columns=1]

\title{{{tex_escape(title)}}} \author{{{tex_escape(latex_author)}}} \date{{\today}}
\begin{{document}}
\maketitle
\tableofcontents
"""
    body = ""
    for i, note in enumerate(notes_info):
        if not Path(note.get("filepath", "")).is_file() or note['pages'] == 0:
            continue
        title_escaped = tex_escape(note["title"])
        d = note['date']
        date_formatted = f"{d[0:4]}/{d[4:6]}/{d[6:8]}" if d.isdigit() and len(d) == 8 else d

        cp_key = note.get('commonplace_key', '')
        icon = key_icons.get(cp_key.lower(), '')
        icon_color_hex = key_colors.get(cp_key.lower())

        icon_latex = ""
        if icon and icon_color_hex:
            rgb_frac = Helper.hex_to_rgb_frac(icon_color_hex)
            icon_latex = fr"\textcolor[rgb]{rgb_frac}{{{icon}}} "
        elif icon:
            icon_latex = f"{icon} "

        header_text = tex_escape(f"{date_formatted} {note['title']}")
        full_header_text = f"{icon_latex}{header_text}"

        body += f"\\multido{{\\i=1+1}}{{{note['pages']}}}{{%\n"
        body += f"  \\ifnum\\i=1\n"
        body += f"    \\clearpage\\phantomsection\n"
        body += f"    \\addcontentsline{{toc}}{{section}}{{{date_formatted} -- {title_escaped}}}\n"
        for tag in note.get("tags", []):
            body += f"    \\index[tags]{{{tex_escape(tag)}!{title_escaped}}}\n"
        if cp_key:
            cp_key_with_icon = f"{icon_latex}{tex_escape(cp_key)}"
            body += f"\\index[cpkeys]{{{cp_key_with_icon}!{title_escaped}}}\n"
        body += f"  \\fi\n"
        body += f"  \\fancyhead[L]{{{full_header_text} (\\i/{note['pages']})}}\n"
        body += f"  \\thispagestyle{{fancy}}\\mbox{{}}\\newpage\n"
        body += f"}}\n"

    postamble = r"""
\clearpage
\fancyhead[L]{{Index Key 索引}}
\printindex[cpkeys]

\clearpage
\fancyhead[L]{{タグ索引}}
\printindex[tags]
\end{document}
"""
    return preamble + body + postamble
