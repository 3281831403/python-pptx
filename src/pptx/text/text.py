"""Text-related objects such as TextFrame and Paragraph."""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, Iterator, cast

from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import MSO_AUTO_SIZE, MSO_UNDERLINE, MSO_VERTICAL_ANCHOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.dml.fill import CT_SolidColorFillProperties
from pptx.oxml.simpletypes import ST_TextWrappingType
from pptx.shapes import Subshape
from pptx.text.fonts import FontFiles
from pptx.text.layout import TextFitter
from pptx.util import Centipoints, Emu, Length, Pt, lazyproperty

if TYPE_CHECKING:
    from pptx.enum.text import (
        MSO_TEXT_UNDERLINE_TYPE,
        MSO_VERTICAL_ANCHOR,
        PP_PARAGRAPH_ALIGNMENT,
    )
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.dml.color import CT_SchemeColor, CT_SRgbColor
    from pptx.oxml.text import (
        CT_RegularTextRun,
        CT_TextBody,
        CT_TextCharacterProperties,
        CT_TextParagraph,
        CT_TextParagraphProperties,
    )
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.types import ProvidesExtents, ProvidesPart


class TextFrame(Subshape):
    """The part of a shape that contains its text.

    Not all shapes have a text frame. Corresponds to the `p:txBody` element that can
    appear as a child element of `p:sp`. Not intended to be constructed directly.
    """

    def __init__(self, txBody: CT_TextBody, parent: ProvidesPart):
        super(TextFrame, self).__init__(parent)
        self._element = self._txBody = txBody
        self._parent = parent

    def add_paragraph(self):
        """
        Return new |_Paragraph| instance appended to the sequence of
        paragraphs contained in this text frame.
        """
        p = self._txBody.add_p()
        return _Paragraph(p, self)

    @property
    def auto_size(self) -> MSO_AUTO_SIZE | None:
        """Resizing strategy used to fit text within this shape.

        Determins the type of automatic resizing used to fit the text of this shape within its
        bounding box when the text would otherwise extend beyond the shape boundaries. May be
        |None|, `MSO_AUTO_SIZE.NONE`, `MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`, or
        `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`.
        """
        return self._bodyPr.autofit

    @auto_size.setter
    def auto_size(self, value: MSO_AUTO_SIZE | None):
        self._bodyPr.autofit = value

    def clear(self):
        """Remove all paragraphs except one empty one."""
        for p in self._txBody.p_lst[1:]:
            self._txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()

    def fit_text(
        self,
        font_family: str = "Calibri",
        max_size: int = 18,
        bold: bool = False,
        italic: bool = False,
        font_file: str | None = None,
    ):
        """Fit text-frame text entirely within bounds of its shape.

        Make the text in this text frame fit entirely within the bounds of its shape by setting
        word wrap on and applying the "best-fit" font size to all the text it contains.

        :attr:`TextFrame.auto_size` is set to :attr:`MSO_AUTO_SIZE.NONE`. The font size will not
        be set larger than `max_size` points. If the path to a matching TrueType font is provided
        as `font_file`, that font file will be used for the font metrics. If `font_file` is |None|,
        best efforts are made to locate a font file with matchhing `font_family`, `bold`, and
        `italic` installed on the current system (usually succeeds if the font is installed).
        """
        # ---no-op when empty as fit behavior not defined for that case---
        if self.text == "":
            return  # pragma: no cover

        font_size = self._best_fit_font_size(font_family, max_size, bold, italic, font_file)
        self._apply_fit(font_family, font_size, bold, italic)

    @property
    def margin_bottom(self) -> Length:
        """|Length| value representing the inset of text from the bottom text frame border.

        :meth:`pptx.util.Inches` provides a convenient way of setting the value, e.g.
        `text_frame.margin_bottom = Inches(0.05)`.
        """
        return self._bodyPr.bIns

    @margin_bottom.setter
    def margin_bottom(self, emu: Length):
        self._bodyPr.bIns = emu

    @property
    def margin_left(self) -> Length:
        """Inset of text from left text frame border as |Length| value."""
        return self._bodyPr.lIns

    @margin_left.setter
    def margin_left(self, emu: Length):
        self._bodyPr.lIns = emu

    @property
    def margin_right(self) -> Length:
        """Inset of text from right text frame border as |Length| value."""
        return self._bodyPr.rIns

    @margin_right.setter
    def margin_right(self, emu: Length):
        self._bodyPr.rIns = emu

    @property
    def margin_top(self) -> Length:
        """Inset of text from top text frame border as |Length| value."""
        return self._bodyPr.tIns

    @margin_top.setter
    def margin_top(self, emu: Length):
        self._bodyPr.tIns = emu

    @property
    def paragraphs(self) -> tuple[_Paragraph, ...]:
        """Sequence of paragraphs in this text frame.

        A text frame always contains at least one paragraph.
        """
        return tuple([_Paragraph(p, self) for p in self._txBody.p_lst])

    @property
    def text(self) -> str:
        """All text in this text-frame as a single string.

        Read/write. The return value contains all text in this text-frame. A line-feed character
        (`"\\n"`) separates the text for each paragraph. A vertical-tab character (`"\\v"`) appears
        for each line break (aka. soft carriage-return) encountered.

        The vertical-tab character is how PowerPoint represents a soft carriage return in clipboard
        text, which is why that encoding was chosen.

        Assignment replaces all text in the text frame. A new paragraph is added for each line-feed
        character (`"\\n"`) encountered. A line-break (soft carriage-return) is inserted for each
        vertical-tab character (`"\\v"`) encountered.

        Each paragraph in the new text is matched with the paragraph at the same position in the
        prior content; that paragraph's properties (such as indentation level) and its explicit
        character formatting are retained, so replaced text keeps the appearance it had. Paragraphs
        beyond the number previously present take their properties from the last prior paragraph.

        Any control character other than newline, tab, or vertical-tab are escaped as plain-text
        like "_x001B_" (for ESC (ASCII 32) in this example).
        """
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @text.setter
    def text(self, text: str):
        txBody = self._txBody
        p_texts = text.split("\n")
        old_ps = list(txBody.p_lst)
        # ---remove paragraphs in excess of those required by the new text---
        for p in old_ps[len(p_texts) :]:
            txBody.remove(p)
        for idx, p_text in enumerate(p_texts):
            if idx < len(old_ps):
                # ---reuse an existing paragraph so its properties are preserved---
                p = old_ps[idx]
            else:
                # ---a further paragraph takes its properties from the last prior one---
                p = txBody.add_p()
                last_pPr = old_ps[-1].pPr if old_ps else None
                if last_pPr is not None:
                    p.insert_element_before(
                        copy.deepcopy(last_pPr), "a:r", "a:br", "a:fld", "a:endParaRPr"
                    )
            _Paragraph(p, self).text = p_text

    @property
    def vertical_anchor(self) -> MSO_VERTICAL_ANCHOR | None:
        """Represents the vertical alignment of text in this text frame.

        |None| indicates the effective value should be inherited from this object's style hierarchy.
        """
        return self._txBody.bodyPr.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, value: MSO_VERTICAL_ANCHOR | None):
        bodyPr = self._txBody.bodyPr
        bodyPr.anchor = value

    @property
    def word_wrap(self) -> bool | None:
        """`True` when lines of text in this shape are wrapped to fit within the shape's width.

        Read-write. Valid values are True, False, or None. True and False turn word wrap on and
        off, respectively. Assigning None to word wrap causes any word wrap setting to be removed
        from the text frame, causing it to inherit this setting from its style hierarchy.
        """
        return {
            ST_TextWrappingType.SQUARE: True,
            ST_TextWrappingType.NONE: False,
            None: None,
        }[self._txBody.bodyPr.wrap]

    @word_wrap.setter
    def word_wrap(self, value: bool | None):
        if value not in (True, False, None):
            raise ValueError(  # pragma: no cover
                "assigned value must be True, False, or None, got %s" % value
            )
        self._txBody.bodyPr.wrap = {
            True: ST_TextWrappingType.SQUARE,
            False: ST_TextWrappingType.NONE,
            None: None,
        }[value]

    def _apply_fit(self, font_family: str, font_size: int, is_bold: bool, is_italic: bool):
        """Arrange text in this text frame to fit inside its extents.

        This is accomplished by setting auto size off, wrap on, and setting the font of
        all its text to `font_family`, `font_size`, `is_bold`, and `is_italic`.
        """
        self.auto_size = MSO_AUTO_SIZE.NONE
        self.word_wrap = True
        self._set_font(font_family, font_size, is_bold, is_italic)

    def _best_fit_font_size(
        self, family: str, max_size: int, bold: bool, italic: bool, font_file: str | None
    ) -> int:
        """Return font-size in points that best fits text in this text-frame.

        The best-fit font size is the largest integer point size not greater than `max_size` that
        allows all the text in this text frame to fit inside its extents when rendered using the
        font described by `family`, `bold`, and `italic`. If `font_file` is specified, it is used
        to calculate the fit, whether or not it matches `family`, `bold`, and `italic`.
        """
        if font_file is None:
            font_file = FontFiles.find(family, bold, italic)
        return TextFitter.best_fit_font_size(self.text, self._extents, max_size, font_file)

    @property
    def _bodyPr(self):
        return self._txBody.bodyPr

    @property
    def _extents(self) -> tuple[Length, Length]:
        """(cx, cy) 2-tuple representing the effective rendering area of this text-frame.

        Margins are taken into account.
        """
        parent = cast("ProvidesExtents", self._parent)
        return (
            Length(parent.width - self.margin_left - self.margin_right),
            Length(parent.height - self.margin_top - self.margin_bottom),
        )

    def _set_font(self, family: str, size: int, bold: bool, italic: bool):
        """Set the font properties of all the text in this text frame."""

        def iter_rPrs(txBody: CT_TextBody) -> Iterator[CT_TextCharacterProperties]:
            for p in txBody.p_lst:
                for elm in p.content_children:
                    yield elm.get_or_add_rPr()
                # ---update an explicit paragraph default when one is present, so the---
                # ---fitted size is also reflected at the paragraph level---
                pPr = p.pPr
                if pPr is not None and pPr.defRPr is not None:
                    yield pPr.defRPr
                # generate a:endParaRPr for each <a:p> element
                yield p.get_or_add_endParaRPr()

        def set_rPr_font(
            rPr: CT_TextCharacterProperties, name: str, size: int, bold: bool, italic: bool
        ):
            f = Font(rPr)
            f.name, f.size, f.bold, f.italic = family, Pt(size), bold, italic

        txBody = self._element
        for rPr in iter_rPrs(txBody):
            set_rPr_font(rPr, family, size, bold, italic)


class Font(object):
    """Character properties object, providing font size, font name, bold, italic, etc.

    Corresponds to `a:rPr` child element of a run. Also appears as `a:defRPr` and
    `a:endParaRPr` in paragraph and `a:defRPr` in list style elements.
    """

    def __init__(self, rPr: CT_TextCharacterProperties):
        super(Font, self).__init__()
        self._element = self._rPr = rPr

    @property
    def bold(self) -> bool | None:
        """Get or set boolean bold value of |Font|, e.g. `paragraph.font.bold = True`.

        If set to |None|, the bold setting is cleared and is inherited from an enclosing shape's
        setting, or a setting in a style or master. Returns None if no bold attribute is present,
        meaning the effective bold value is inherited from a master or the theme.
        """
        return self._rPr.b

    @bold.setter
    def bold(self, value: bool | None):
        self._rPr.b = value

    @lazyproperty
    def color(self) -> ColorFormat:
        """The |ColorFormat| instance that provides access to the color settings for this font.

        Accessing this property never changes the XML. When no fill is specified on the run
        properties, the returned object reports a color type of |None| and an `a:solidFill`
        element is only added if and when a color value is assigned. When the font has a
        non-solid fill, such as a gradient, a |TypeError| is raised; use :attr:`fill` to
        access the fill in that case.
        """
        if self._rPr.eg_fillProperties is None:
            return ColorFormat.from_colorchoice_parent(_DeferredSolidFill(self._rPr))
        return self.fill.fore_color

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this font.

        Provides access to fill properties such as fill color.
        """
        return FillFormat.from_fill_parent(self._rPr)

    @property
    def italic(self) -> bool | None:
        """Get or set boolean italic value of |Font| instance.

        Has the same behaviors as bold with respect to None values.
        """
        return self._rPr.i

    @italic.setter
    def italic(self, value: bool | None):
        self._rPr.i = value

    @property
    def language_id(self) -> MSO_LANGUAGE_ID | None:
        """Get or set the language id of this |Font| instance.

        The language id is a member of the :ref:`MsoLanguageId` enumeration. Assigning |None|
        removes any language setting, the same behavior as assigning `MSO_LANGUAGE_ID.NONE`.
        """
        lang = self._rPr.lang
        if lang is None:
            return MSO_LANGUAGE_ID.NONE
        return self._rPr.lang

    @language_id.setter
    def language_id(self, value: MSO_LANGUAGE_ID | None):
        if value == MSO_LANGUAGE_ID.NONE:
            value = None
        self._rPr.lang = value

    @property
    def name(self) -> str | None:
        """Get or set the typeface name for this |Font| instance.

        Causes the text it controls to appear in the named font, if a matching font is found.
        Returns |None| if the typeface is currently inherited from the theme. Setting it to |None|
        removes any override of the theme typeface.
        """
        latin = self._rPr.latin
        if latin is None:
            return None
        return latin.typeface

    @name.setter
    def name(self, value: str | None):
        if value is None:
            self._rPr._remove_latin()  # pyright: ignore[reportPrivateUsage]
        else:
            latin = self._rPr.get_or_add_latin()
            latin.typeface = value

    @property
    def size(self) -> Length | None:
        """Indicates the font height in English Metric Units (EMU).

        Read/write. |None| indicates the font size should be inherited from its style hierarchy,
        such as a placeholder or document defaults (usually 18pt). |Length| is a subclass of |int|
        having properties for convenient conversion into points or other length units. Likewise,
        the :class:`pptx.util.Pt` class allows convenient specification of point values::

            >>> font.size = Pt(24)
            >>> font.size
            304800
            >>> font.size.pt
            24.0
        """
        sz = self._rPr.sz
        if sz is None:
            return None
        return Centipoints(sz)

    @size.setter
    def size(self, emu: Length | None):
        if emu is None:
            self._rPr.sz = None
        else:
            sz = Emu(emu).centipoints
            self._rPr.sz = sz

    @property
    def underline(self) -> bool | MSO_TEXT_UNDERLINE_TYPE | None:
        """Indicaties the underline setting for this font.

        Value is |True|, |False|, |None|, or a member of the :ref:`MsoTextUnderlineType`
        enumeration. |None| is the default and indicates the underline setting should be inherited
        from the style hierarchy, such as from a placeholder. |True| indicates single underline.
        |False| indicates no underline. Other settings such as double and wavy underlining are
        indicated with members of the :ref:`MsoTextUnderlineType` enumeration.
        """
        u = self._rPr.u
        if u is MSO_UNDERLINE.NONE:
            return False
        if u is MSO_UNDERLINE.SINGLE_LINE:
            return True
        return u

    @underline.setter
    def underline(self, value: bool | MSO_TEXT_UNDERLINE_TYPE | None):
        if value is True:
            value = MSO_UNDERLINE.SINGLE_LINE
        elif value is False:
            value = MSO_UNDERLINE.NONE
        self._element.u = value


class _DeferredSolidFill:
    """`EG_ColorChoice` parent that stands in for an `a:solidFill` element not yet present.

    Allows the font color to be read when no fill is specified without adding any elements
    to the XML. The `a:solidFill` element is materialized on *fill_parent* (an element
    having the `EG_FillProperties` child group, or a proxy for one) only when a color value
    is assigned.
    """

    def __init__(self, fill_parent: CT_TextCharacterProperties | _ParagraphFillParent):
        super(_DeferredSolidFill, self).__init__()
        self._fill_parent = fill_parent

    @property
    def eg_colorChoice(self) -> BaseOxmlElement | None:
        """Color-choice child of the solid fill, None when no solid fill is present."""
        solidFill = self._solidFill
        if solidFill is None:
            return None
        return cast("BaseOxmlElement | None", solidFill.eg_colorChoice)

    def get_or_change_to_schemeClr(self) -> CT_SchemeColor:
        """Return the `a:schemeClr` child of the solid fill, added as needed."""
        return self._get_or_add_solidFill().get_or_change_to_schemeClr()

    def get_or_change_to_srgbClr(self) -> CT_SRgbColor:
        """Return the `a:srgbClr` child of the solid fill, added as needed."""
        return self._get_or_add_solidFill().get_or_change_to_srgbClr()

    def _get_or_add_solidFill(self) -> CT_SolidColorFillProperties:
        """The `a:solidFill` element of the fill parent, materialized if not present."""
        solidFill = self._solidFill
        if solidFill is not None:
            return solidFill
        return self._fill_parent.get_or_change_to_solidFill()

    @property
    def _solidFill(self) -> CT_SolidColorFillProperties | None:
        """The `a:solidFill` child of the fill parent, None when a different or no fill is set."""
        fill = self._fill_parent.eg_fillProperties
        if isinstance(fill, CT_SolidColorFillProperties):
            return fill
        return None


class _ParagraphFillParent:
    """`EG_FillProperties` parent view over a paragraph's explicit character properties.

    A fill is read from `a:pPr/a:defRPr` when one is specified there, otherwise from
    `a:endParaRPr`; no elements are added to the XML for reading. A fill-changing method
    applies the change to `a:defRPr` (added as needed) and removes any fill from
    `a:endParaRPr`, so the two cannot hold contradicting values.
    """

    def __init__(self, p: CT_TextParagraph):
        super(_ParagraphFillParent, self).__init__()
        self._p = p

    @property
    def eg_fillProperties(self) -> BaseOxmlElement | None:
        """Fill element of `a:defRPr`, or of `a:endParaRPr` when defRPr specifies none."""
        for rPr in (self._defRPr, self._p.endParaRPr):
            if rPr is not None:
                fill = rPr.eg_fillProperties
                if fill is not None:
                    return fill
        return None

    def get_or_change_to_blipFill(self) -> BaseOxmlElement:
        return self._change_fill_to("blipFill")

    def get_or_change_to_gradFill(self) -> BaseOxmlElement:
        return self._change_fill_to("gradFill")

    def get_or_change_to_grpFill(self) -> BaseOxmlElement:
        return self._change_fill_to("grpFill")

    def get_or_change_to_noFill(self) -> BaseOxmlElement:
        return self._change_fill_to("noFill")

    def get_or_change_to_pattFill(self) -> BaseOxmlElement:
        return self._change_fill_to("pattFill")

    def get_or_change_to_solidFill(self) -> CT_SolidColorFillProperties:
        return cast(CT_SolidColorFillProperties, self._change_fill_to("solidFill"))

    def _change_fill_to(self, fill_tag: str) -> BaseOxmlElement:
        """Return the `a:{fill_tag}` child of `a:defRPr`, added or changed-to as needed.

        Any fill on `a:endParaRPr` is removed so it cannot contradict the new value.
        """
        defRPr = self._get_or_add_defRPr()
        endParaRPr = self._p.endParaRPr
        if endParaRPr is not None:
            endParaRPr._remove_eg_fillProperties()  # pyright: ignore[reportPrivateUsage]
        change_to = {
            "blipFill": defRPr.get_or_change_to_blipFill,
            "gradFill": defRPr.get_or_change_to_gradFill,
            "grpFill": defRPr.get_or_change_to_grpFill,
            "noFill": defRPr.get_or_change_to_noFill,
            "pattFill": defRPr.get_or_change_to_pattFill,
            "solidFill": defRPr.get_or_change_to_solidFill,
        }[fill_tag]
        return change_to()

    @property
    def _defRPr(self) -> CT_TextCharacterProperties | None:
        """The `a:defRPr` child of this paragraph's `a:pPr`, None when not present."""
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.defRPr

    def _get_or_add_defRPr(self) -> CT_TextCharacterProperties:
        """The `a:defRPr` element for this paragraph, added with its ancestors as needed."""
        return self._p.get_or_add_pPr().get_or_add_defRPr()


class _ParagraphFont(Font):
    """|Font| subclass bound to the default character properties of a paragraph.

    Values are read from `a:pPr/a:defRPr` when present, falling back to `a:endParaRPr`
    (character formatting applied to the paragraph mark). Reading never adds elements to
    the XML. Writes go to `a:defRPr` (added as needed) and any contradicting value on
    `a:endParaRPr` is removed, so the paragraph properties remain authoritative.
    """

    def __init__(self, p: CT_TextParagraph):
        # ---NOTE: Font.__init__() is intentionally skipped; this subclass is bound to---
        # ---the paragraph element rather than to a concrete character-properties element---
        super(Font, self).__init__()
        self._p = p
        self._fill_parent = _ParagraphFillParent(p)

    @property
    def bold(self) -> bool | None:
        """Explicit bold value for this paragraph, None when inherited.

        Read from `a:defRPr`, falling back to `a:endParaRPr`. Assigning |None| removes any
        explicit setting, causing the effective value to be inherited.
        """
        return self._explicit("b")

    @bold.setter
    def bold(self, value: bool | None):
        self._set_explicit("b", value)

    @lazyproperty
    def color(self) -> ColorFormat:
        """The |ColorFormat| instance providing access to the color settings for this font.

        Accessing this property never changes the XML. When no fill is specified, the
        returned object reports a color type of |None| and an `a:solidFill` element is only
        added (to `a:defRPr`) if and when a color value is assigned. When the font has a
        non-solid fill, such as a gradient, a |TypeError| is raised; use :attr:`fill` to
        access the fill in that case.
        """
        if self._fill_parent.eg_fillProperties is None:
            return ColorFormat.from_colorchoice_parent(_DeferredSolidFill(self._fill_parent))
        return self.fill.fore_color

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this font, providing access to fill properties.

        Reflects the fill of `a:defRPr`, or of `a:endParaRPr` when defRPr specifies none.
        Changes are applied to `a:defRPr` and any fill on `a:endParaRPr` is removed.
        """
        return FillFormat.from_fill_parent(cast("BaseOxmlElement", self._fill_parent))

    @property
    def italic(self) -> bool | None:
        """Explicit italic value for this paragraph, with the same behaviors as :attr:`bold`."""
        return self._explicit("i")

    @italic.setter
    def italic(self, value: bool | None):
        self._set_explicit("i", value)

    @property
    def language_id(self) -> MSO_LANGUAGE_ID | None:
        """Explicit language id for this paragraph, read from defRPr or endParaRPr.

        Assigning |None| or `MSO_LANGUAGE_ID.NONE` removes any language setting.
        """
        lang = self._explicit("lang")
        if lang is None:
            return MSO_LANGUAGE_ID.NONE
        return lang

    @language_id.setter
    def language_id(self, value: MSO_LANGUAGE_ID | None):
        if value == MSO_LANGUAGE_ID.NONE:
            value = None
        self._set_explicit("lang", value)

    @property
    def name(self) -> str | None:
        """Explicit typeface name for this paragraph, read from defRPr or endParaRPr.

        Assigning |None| removes any override of the inherited typeface.
        """
        for rPr in (self._defRPr, self._p.endParaRPr):
            if rPr is not None and rPr.latin is not None:
                return rPr.latin.typeface
        return None

    @name.setter
    def name(self, value: str | None):
        if value is None:
            for rPr in (self._defRPr, self._p.endParaRPr):
                if rPr is not None and rPr.latin is not None:
                    rPr._remove_latin()  # pyright: ignore[reportPrivateUsage]
            return
        latin = self._get_or_add_defRPr().get_or_add_latin()
        latin.typeface = value
        endParaRPr = self._p.endParaRPr
        if endParaRPr is not None and endParaRPr.latin is not None:
            endParaRPr._remove_latin()  # pyright: ignore[reportPrivateUsage]

    @property
    def size(self) -> Length | None:
        """Explicit font size for this paragraph, read from defRPr or endParaRPr.

        |None| indicates the font size is inherited from the style hierarchy. Assigning
        |None| removes any explicit size setting.
        """
        sz = self._explicit("sz")
        if sz is None:
            return None
        return Centipoints(sz)

    @size.setter
    def size(self, emu: Length | None):
        sz = None if emu is None else Emu(emu).centipoints
        self._set_explicit("sz", sz)

    @property
    def underline(self) -> bool | MSO_TEXT_UNDERLINE_TYPE | None:
        """Explicit underline setting for this paragraph, read from defRPr or endParaRPr."""
        u = self._explicit("u")
        if u is MSO_UNDERLINE.NONE:
            return False
        if u is MSO_UNDERLINE.SINGLE_LINE:
            return True
        return u

    @underline.setter
    def underline(self, value: bool | MSO_TEXT_UNDERLINE_TYPE | None):
        if value is True:
            value = MSO_UNDERLINE.SINGLE_LINE
        elif value is False:
            value = MSO_UNDERLINE.NONE
        self._set_explicit("u", value)

    def _explicit(self, attr_name: str):
        """Value of `attr_name` from `a:defRPr`, falling back to `a:endParaRPr`.

        None when neither element specifies the attribute. Never adds elements to the XML.
        """
        for rPr in (self._defRPr, self._p.endParaRPr):
            if rPr is not None:
                value = getattr(rPr, attr_name)
                if value is not None:
                    return value
        return None

    def _set_explicit(self, attr_name: str, value: object) -> None:
        """Write `value` to `a:defRPr` and remove any contradicting value from `a:endParaRPr`.

        `a:defRPr` is added as needed, except when `value` is None, which simply removes
        any explicit setting without adding elements.
        """
        if value is None:
            for rPr in (self._defRPr, self._p.endParaRPr):
                if rPr is not None:
                    setattr(rPr, attr_name, None)
            return
        setattr(self._get_or_add_defRPr(), attr_name, value)
        endParaRPr = self._p.endParaRPr
        if endParaRPr is not None:
            setattr(endParaRPr, attr_name, None)

    @property
    def _defRPr(self) -> CT_TextCharacterProperties | None:
        """The `a:defRPr` child of this paragraph's `a:pPr`, None when not present."""
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.defRPr

    def _get_or_add_defRPr(self) -> CT_TextCharacterProperties:
        """The `a:defRPr` element for this paragraph, added with its ancestors as needed."""
        return self._p.get_or_add_pPr().get_or_add_defRPr()


class _Hyperlink(Subshape):
    """Text run hyperlink object.

    Corresponds to `a:hlinkClick` child element of the run's properties element (`a:rPr`).
    """

    def __init__(self, rPr: CT_TextCharacterProperties, parent: ProvidesPart):
        super(_Hyperlink, self).__init__(parent)
        self._rPr = rPr

    @property
    def address(self) -> str | None:
        """The URL of the hyperlink.

        Read/write. URL can be on http, https, mailto, or file scheme; others may work.
        """
        if self._hlinkClick is None:
            return None
        return self.part.target_ref(self._hlinkClick.rId)

    @address.setter
    def address(self, url: str | None):
        # implements all three of add, change, and remove hyperlink
        if self._hlinkClick is not None:
            self._remove_hlinkClick()
        if url:
            self._add_hlinkClick(url)

    def _add_hlinkClick(self, url: str):
        rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)
        self._rPr.add_hlinkClick(rId)

    @property
    def _hlinkClick(self) -> CT_Hyperlink | None:
        return self._rPr.hlinkClick

    def _remove_hlinkClick(self):
        assert self._hlinkClick is not None
        self.part.drop_rel(self._hlinkClick.rId)
        self._rPr._remove_hlinkClick()  # pyright: ignore[reportPrivateUsage]


class _Paragraph(Subshape):
    """Paragraph object. Not intended to be constructed directly."""

    def __init__(self, p: CT_TextParagraph, parent: ProvidesPart):
        super(_Paragraph, self).__init__(parent)
        self._element = self._p = p

    def add_line_break(self):
        """Add line break at end of this paragraph."""
        self._p.add_br()

    def add_run(self) -> _Run:
        """Return a new run appended to the runs in this paragraph."""
        r = self._p.add_r()
        return _Run(r, self)

    @property
    def alignment(self) -> PP_PARAGRAPH_ALIGNMENT | None:
        """Horizontal alignment of this paragraph.

        The value |None| indicates the paragraph should 'inherit' its effective value from its
        style hierarchy. Assigning |None| removes any explicit setting, causing its inherited
        value to be used.
        """
        return self._pPr.algn

    @alignment.setter
    def alignment(self, value: PP_PARAGRAPH_ALIGNMENT | None):
        self._pPr.algn = value

    def clear(self):
        """Remove all content from this paragraph.

        Paragraph properties are preserved. Content includes runs, line breaks, and fields.
        """
        for elm in self._element.content_children:
            self._element.remove(elm)
        return self

    @property
    def font(self) -> Font:
        """|Font| object containing default character properties for the runs in this paragraph.

        These character properties override default properties inherited from parent objects such
        as the text frame the paragraph is contained in and they may be overridden by character
        properties set at the run level.

        Values are read from `a:pPr/a:defRPr` when present, falling back to `a:endParaRPr`
        (formatting applied to the paragraph mark); reading never adds elements to the XML.
        Assigned values are written to `a:defRPr` and any contradicting value on `a:endParaRPr`
        is removed.
        """
        return _ParagraphFont(self._p)

    @property
    def level(self) -> int:
        """Indentation level of this paragraph.

        Read-write. Integer in range 0..8 inclusive. 0 represents a top-level paragraph and is the
        default value. Indentation level is most commonly encountered in a bulleted list, as is
        found on a word bullet slide.
        """
        return self._pPr.lvl

    @level.setter
    def level(self, level: int):
        self._pPr.lvl = level

    @property
    def line_spacing(self) -> int | float | Length | None:
        """The space between baselines in successive lines of this paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. A numeric value, e.g. `2` or `1.5`,
        indicates spacing is applied in multiples of line heights. A |Length| value such as
        `Pt(12)` indicates spacing is a fixed height. The |Pt| value class is a convenient way to
        apply line spacing in units of points.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.line_spacing

    @line_spacing.setter
    def line_spacing(self, value: int | float | Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.line_spacing = value

    @property
    def runs(self) -> tuple[_Run, ...]:
        """Sequence of runs in this paragraph."""
        return tuple(_Run(r, self) for r in self._element.r_lst)

    @property
    def space_after(self) -> Length | None:
        """The spacing to appear between this paragraph and the subsequent paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. |Length| objects provide convenience
        properties, such as `.pt` and `.inches`, that allow easy conversion to various length
        units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_after

    @space_after.setter
    def space_after(self, value: Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.space_after = value

    @property
    def space_before(self) -> Length | None:
        """The spacing to appear between this paragraph and the prior paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. |Length| objects provide convenience
        properties, such as `.pt` and `.cm`, that allow easy conversion to various length units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_before

    @space_before.setter
    def space_before(self, value: Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.space_before = value

    @property
    def text(self) -> str:
        """Text of paragraph as a single string.

        Read/write. This value is formed by concatenating the text in each run and field making up
        the paragraph, adding a vertical-tab character (`"\\v"`) for each line-break element
        (`<a:br>`, soft carriage-return) encountered.

        While the encoding of line-breaks as a vertical tab might be surprising at first, doing so
        is consistent with PowerPoint's clipboard copy behavior and allows a line-break to be
        distinguished from a paragraph boundary within the str return value.

        Assignment causes all content in the paragraph to be replaced. Each vertical-tab character
        (`"\\v"`) in the assigned str is translated to a line-break, as is each line-feed
        character (`"\\n"`). Contrast behavior of line-feed character in `TextFrame.text` setter.
        If line-feed characters are intended to produce new paragraphs, use `TextFrame.text`
        instead. Any other control characters in the assigned string are escaped as a hex
        representation like "_x001B_" (for ESC (ASCII 27) in this example).

        The replacement text inherits this paragraph's explicit character formatting, so
        settings such as font size, bold, and color are retained when text is replaced.
        """
        return "".join(elm.text for elm in self._element.content_children)

    @text.setter
    def text(self, text: str):
        self._promote_explicit_char_format()
        self.clear()
        self._element.append_text(text)

    @property
    def _pPr(self) -> CT_TextParagraphProperties:
        """Contains the properties for this paragraph.

        Causes the element to be added if not present.
        """
        return self._p.get_or_add_pPr()

    def _promote_explicit_char_format(self):
        """Copy this paragraph's explicit character formatting to `a:pPr/a:defRPr` when needed.

        Runs added to replace this paragraph's content inherit their character formatting
        from `a:defRPr`. When this paragraph carries explicit character formatting only on
        its paragraph mark (`a:endParaRPr`) or on its first content item (run, line-break,
        or field), that formatting is copied to `a:defRPr` so replacement text retains it.
        Does nothing when `a:defRPr` is already present or the paragraph has no explicit
        character formatting; in particular no elements are added for a paragraph having
        no explicit formatting.
        """
        p = self._p
        pPr = p.pPr
        if pPr is not None and pPr.defRPr is not None:
            return
        source: CT_TextCharacterProperties | None = p.endParaRPr
        if source is None:
            content = p.content_children
            source = content[0].rPr if content else None
        if source is None or (not source.attrib and len(source) == 0):
            return
        defRPr = p.get_or_add_pPr().get_or_add_defRPr()
        for name, value in source.attrib.items():
            defRPr.set(name, value)
        for child in source:
            defRPr.append(copy.deepcopy(child))


class _Run(Subshape):
    """Text run object. Corresponds to `a:r` child element in a paragraph."""

    def __init__(self, r: CT_RegularTextRun, parent: ProvidesPart):
        super(_Run, self).__init__(parent)
        self._r = r

    @property
    def font(self):
        """|Font| instance containing run-level character properties for the text in this run.

        Character properties can be and perhaps most often are inherited from parent objects such
        as the paragraph and slide layout the run is contained in. Only those specifically
        overridden at the run level are contained in the font object.
        """
        rPr = self._r.get_or_add_rPr()
        return Font(rPr)

    @lazyproperty
    def hyperlink(self) -> _Hyperlink:
        """Proxy for any `a:hlinkClick` element under the run properties element.

        Created on demand, the hyperlink object is available whether an `a:hlinkClick` element is
        present or not, and creates or deletes that element as appropriate in response to actions
        on its methods and attributes.
        """
        rPr = self._r.get_or_add_rPr()
        return _Hyperlink(rPr, self)

    @property
    def text(self):
        """Read/write. A unicode string containing the text in this run.

        Assignment replaces all text in the run. The assigned value can be a 7-bit ASCII
        string, a UTF-8 encoded 8-bit string, or unicode. String values are converted to
        unicode assuming UTF-8 encoding.

        Any other control characters in the assigned string other than tab or newline
        are escaped as a hex representation. For example, ESC (ASCII 27) is escaped as
        "_x001B_". Contrast the behavior of `TextFrame.text` and `_Paragraph.text` with
        respect to line-feed and vertical-tab characters.
        """
        return self._r.text

    @text.setter
    def text(self, text: str):
        self._r.text = text
