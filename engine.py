# -*- coding: utf-8 -*-
"""WordSoft engine — نموذج المستند مستقل عن الواجهة (مثل wordcraft-doc).
يدعم: فقرات + Runs بتنسيق، عناوين، قوائم، جداول، صور (مسارات)، تعداد كلمات.
"""
from __future__ import annotations
import re
import os
from dataclasses import dataclass, field, asdict


@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    color: str | None = None        # hex مثل FF0000
    highlight: str | None = None    # hex خلفية
    font: str | None = None
    size: int | None = None         # pt


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    style: str = "Normal"           # Normal, Title, Heading 1/2/3, Quote, List Bullet, List Number
    align: str = "left"             # left, center, right, justify
    direction: str = "auto"         # auto, rtl, ltr
    spacing_before: int = 0
    spacing_after: int = 6
    line_spacing: float = 1.15
    indent: int = 0

    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class Table:
    rows: list[list[str]] = field(default_factory=list)
    style: str = "Table Grid"
    header: bool = True


Block = Paragraph | Table


@dataclass
class WordDoc:
    title: str = "مستند بدون عنوان"
    author: str = "WordSoft"
    blocks: list = field(default_factory=list)
    track_changes: bool = False

    # ---------- إنشاء ----------
    @classmethod
    def blank(cls, title="مستند بدون عنوان") -> "WordDoc":
        d = cls(title=title)
        d.blocks.append(Paragraph([Run("")], style="Normal"))
        return d

    @classmethod
    def sample(cls) -> "WordDoc":
        """مستند عينة عربي مثل The Open Studio Handbook في WordCraft."""
        d = cls(title="دليل الاستوديو المفتوح — WordSoft")
        d.blocks.append(Paragraph([Run("دليل الاستوديو المفتوح", bold=True, size=28)], style="Title", align="center"))
        d.blocks.append(Paragraph([Run("اكتب وصمّم مستنداتك — معالج نصوص عربي حر مستوحى من WordCraft", size=13)], style="Subtitle", align="center"))
        d.blocks.append(Paragraph([Run("مقدمة", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("مرحباً بك في ") , Run("WordSoft", bold=True), Run(" — معالج النصوص العربي الخفيف الذي يعمل فوراً: تبويبات الشريط، الأنماط، الجداول، تعداد الكلمات، والتصدير إلى DOCX و PDF. جرّب تحديد هذا السطر واضغط "), Run("Ctrl+B", bold=True), Run(" للعريض.")]))
        d.blocks.append(Paragraph([Run("الأنماط", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("استخدم معرض الأنماط من تبويب الرئيسية: عنوان، عنوان 1، عنوان 2، اقتباس، وقوائم نقطية ومرقمة. كل نمط يظهر بمعاينة حية.")], style="Normal"))
        d.blocks.append(Paragraph([Run("عنوان فرعي تجريبي", bold=True, size=15)], style="Heading 2"))
        d.blocks.append(Paragraph([Run("هذا اقتباس تجريبي يوضح نمط الاقتباس في WordSoft.", italic=True)], style="Quote"))
        d.blocks.append(Paragraph([Run("نقطة أولى: الكتابة السريعة مع تراجع وإعادة")], style="List Bullet"))
        d.blocks.append(Paragraph([Run("نقطة ثانية: فتح وحفظ DOCX يعمل مع Word")], style="List Bullet"))
        d.blocks.append(Paragraph([Run("خطوة أولى: افتح تبويب إدراج وأضف جدولاً")], style="List Number"))
        d.blocks.append(Paragraph([Run("خطوة ثانية: صدّر إلى PDF من قائمة ملف")], style="List Number"))
        d.blocks.append(Paragraph([Run("جدول تجريبي", bold=True, size=18)], style="Heading 1"))
        t = Table(rows=[["العنصر", "الوصف", "الحالة"], ["الأنماط", "عناوين + اقتباس + قوائم", "يعمل"], ["الجداول", "إدراج وتحرير وتصدير", "يعمل"], ["PDF", "تصدير حقيقي بنص قابل للتحديد", "يعمل"]], header=True)
        d.blocks.append(t)
        d.blocks.append(Paragraph([Run("خاتمة", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("WordSoft مستوحى من فكرة WordCraft (إعادة بناء Word مفتوحة المصدر) لكن مكتوباً بلغة Python ليعمل فوراً بدون تجميع. كل الأوامر معرفة بمعرفات مثل WordCraft لتكون قابلة للأتمتة عبر CLI.")]))
        return d

    @classmethod
    def template(cls, kind: str) -> "WordDoc":
        if kind == "sample":
            return cls.sample()
        d = cls(title={"letter": "رسالة رسمية", "report": "تقرير", "resume": "سيرة ذاتية"}.get(kind, "مستند بدون عنوان"))
        if kind == "letter":
            d.blocks.append(Paragraph([Run("الاسم الكامل", bold=True, size=16)], align="right"))
            d.blocks.append(Paragraph([Run("التاريخ: .... / .... / ......")], align="right"))
            d.blocks.append(Paragraph([Run("الموضوع: ....................", bold=True)], align="right"))
            d.blocks.append(Paragraph([Run("السلام عليكم ورحمة الله وبركاته،")], align="right"))
            d.blocks.append(Paragraph([Run("أما بعد، يشرفني أن أتقدم إليكم بـ...")], align="right"))
            d.blocks.append(Paragraph([Run("وتقبلوا فائق الاحترام والتقدير.")], align="right"))
            d.blocks.append(Paragraph([Run("التوقيع", italic=True)], align="right"))
        elif kind == "report":
            d.blocks.append(Paragraph([Run("تقرير", bold=True, size=26)], style="Title", align="center"))
            d.blocks.append(Paragraph([Run("ملخص تنفيذي", bold=True, size=18)], style="Heading 1"))
            d.blocks.append(Paragraph([Run("اكتب الملخص هنا...")]))
            d.blocks.append(Paragraph([Run("النتائج", bold=True, size=18)], style="Heading 1"))
            d.blocks.append(Paragraph([Run("النتيجة الأولى...")], style="List Bullet"))
            d.blocks.append(Paragraph([Run("التوصيات", bold=True, size=18)], style="Heading 1"))
        elif kind == "resume":
            d.blocks.append(Paragraph([Run("الاسم الكامل", bold=True, size=24)], style="Title", align="center"))
            d.blocks.append(Paragraph([Run("البريد | الهاتف | العنوان")], align="center"))
            d.blocks.append(Paragraph([Run("الخبرات", bold=True, size=18)], style="Heading 1"))
            d.blocks.append(Paragraph([Run("المسمى الوظيفي — الشركة (السنة)")], style="List Bullet"))
            d.blocks.append(Paragraph([Run("المهارات", bold=True, size=18)], style="Heading 1"))
        else:
            return cls.blank()
        return d

    # ---------- إحصاء ----------
    def full_text(self) -> str:
        out = []
        for b in self.blocks:
            if isinstance(b, Paragraph):
                out.append(b.text())
            elif isinstance(b, Table):
                for row in b.rows:
                    out.append(" | ".join(row))
        return "\n".join(out)

    def stats(self) -> dict:
        text = self.full_text()
        words = re.findall(r"\S+", text)
        chars = len(text.replace("\n", ""))
        paras = sum(1 for b in self.blocks if isinstance(b, Paragraph) and b.text().strip())
        pages = max(1, round(len(words) / 400 + 0.5))
        return {"words": len(words), "chars": chars, "paragraphs": paras,
                "tables": sum(1 for b in self.blocks if isinstance(b, Table)),
                "pages": pages}

    def headings(self) -> list[tuple[str, str]]:
        """للتنقل: (style, text) للعناوين فقط."""
        out = []
        for b in self.blocks:
            if isinstance(b, Paragraph) and b.style.startswith("Heading") and b.text().strip():
                out.append((b.style, b.text().strip()))
        return out
