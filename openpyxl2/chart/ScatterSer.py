#Autogenerated schema
        from openpyxl2.descriptors.serialisable import Serialisable
        from openpyxl2.descriptors import (
    Typed,
    Integer,
    Float,
    NoneSet,
    String,
    Set,
    Bool,)


class AxDataSource(Serialisable):

    pass

class NumDataSource(Serialisable):

    pass

class ErrValType(Serialisable):

    val = Set(values=(['cust', 'fixedVal', 'percentage', 'stdDev', 'stdErr']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ErrBarType(Serialisable):

    val = Set(values=(['both', 'minus', 'plus']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ErrDir(Serialisable):

    val = Set(values=(['x', 'y']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ErrBars(Serialisable):

    errDir = Typed(expected_type=ErrDir, allow_none=True)
    errBarType = Typed(expected_type=ErrBarType, )
    errValType = Typed(expected_type=ErrValType, )
    noEndCap = Bool(nested=True, allow_none=True)
    plus = Typed(expected_type=NumDataSource, allow_none=True)
    minus = Typed(expected_type=NumDataSource, allow_none=True)
    val = Float(nested=True, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 errDir=None,
                 errBarType=None,
                 errValType=None,
                 noEndCap=None,
                 plus=None,
                 minus=None,
                 val=None,
                 spPr=None,
                 extLst=None,
                ):
        self.errDir = errDir
        self.errBarType = errBarType
        self.errValType = errValType
        self.noEndCap = noEndCap
        self.plus = plus
        self.minus = minus
        self.val = val
        self.spPr = spPr
        self.extLst = extLst


class TextParagraph(Serialisable):

    pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    endParaRPr = Typed(expected_type=TextCharacterProperties, allow_none=True)

    def __init__(self,
                 pPr=None,
                 endParaRPr=None,
                ):
        self.pPr = pPr
        self.endParaRPr = endParaRPr


class EmbeddedWAVAudioFile(Serialisable):

    name = String(allow_none=True)

    def __init__(self,
                 name=None,
                ):
        self.name = name


class Hyperlink(Serialisable):

    invalidUrl = String(allow_none=True)
    action = String(allow_none=True)
    tgtFrame = String(allow_none=True)
    tooltip = String(allow_none=True)
    history = Bool(allow_none=True)
    highlightClick = Bool(allow_none=True)
    endSnd = Bool(allow_none=True)
    snd = Typed(expected_type=EmbeddedWAVAudioFile, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 invalidUrl=None,
                 action=None,
                 tgtFrame=None,
                 tooltip=None,
                 history=None,
                 highlightClick=None,
                 endSnd=None,
                 snd=None,
                 extLst=None,
                ):
        self.invalidUrl = invalidUrl
        self.action = action
        self.tgtFrame = tgtFrame
        self.tooltip = tooltip
        self.history = history
        self.highlightClick = highlightClick
        self.endSnd = endSnd
        self.snd = snd
        self.extLst = extLst


class TextFont(Serialisable):

    typeface = String()
    panose = HexBinary(allow_none=True)
    pitchFamily = Integer(allow_none=True)
    charset = Integer(allow_none=True)

    def __init__(self,
                 typeface=None,
                 panose=None,
                 pitchFamily=None,
                 charset=None,
                ):
        self.typeface = typeface
        self.panose = panose
        self.pitchFamily = pitchFamily
        self.charset = charset


class TextCharacterProperties(Serialisable):

    kumimoji = Bool(allow_none=True)
    lang = String(allow_none=True)
    altLang = String(allow_none=True)
    sz = Integer()
    b = Bool(allow_none=True)
    i = Bool(allow_none=True)
    u = NoneSet(values=(['words', 'sng', 'dbl', 'heavy', 'dotted', 'dottedHeavy', 'dash', 'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy', 'dotDotDash', 'dotDotDashHeavy', 'wavy', 'wavyHeavy', 'wavyDbl']))
    strike = Set(values=(['noStrike', 'sngStrike', 'dblStrike']))
    kern = Integer()
    cap = NoneSet(values=(['small', 'all']))
    spc = unknown defintion for ST_TextPoint
    normalizeH = Bool(allow_none=True)
    baseline = unknown defintion for ST_Percentage
    noProof = Bool(allow_none=True)
    dirty = Bool(allow_none=True)
    err = Bool(allow_none=True)
    smtClean = Bool(allow_none=True)
    smtId = Integer(allow_none=True)
    bmk = String(allow_none=True)
    ln = Typed(expected_type=LineProperties, allow_none=True)
    highlight = Typed(expected_type=Color, allow_none=True)
    latin = Typed(expected_type=TextFont, allow_none=True)
    ea = Typed(expected_type=TextFont, allow_none=True)
    cs = Typed(expected_type=TextFont, allow_none=True)
    sym = Typed(expected_type=TextFont, allow_none=True)
    hlinkClick = Typed(expected_type=Hyperlink, allow_none=True)
    hlinkMouseOver = Typed(expected_type=Hyperlink, allow_none=True)
    rtl = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 kumimoji=None,
                 lang=None,
                 altLang=None,
                 sz=None,
                 b=None,
                 i=None,
                 u=None,
                 strike=None,
                 kern=None,
                 cap=None,
                 spc=None,
                 normalizeH=None,
                 baseline=None,
                 noProof=None,
                 dirty=None,
                 err=None,
                 smtClean=None,
                 smtId=None,
                 bmk=None,
                 ln=None,
                 highlight=None,
                 latin=None,
                 ea=None,
                 cs=None,
                 sym=None,
                 hlinkClick=None,
                 hlinkMouseOver=None,
                 rtl=None,
                 extLst=None,
                ):
        self.kumimoji = kumimoji
        self.lang = lang
        self.altLang = altLang
        self.sz = sz
        self.b = b
        self.i = i
        self.u = u
        self.strike = strike
        self.kern = kern
        self.cap = cap
        self.spc = spc
        self.normalizeH = normalizeH
        self.baseline = baseline
        self.noProof = noProof
        self.dirty = dirty
        self.err = err
        self.smtClean = smtClean
        self.smtId = smtId
        self.bmk = bmk
        self.ln = ln
        self.highlight = highlight
        self.latin = latin
        self.ea = ea
        self.cs = cs
        self.sym = sym
        self.hlinkClick = hlinkClick
        self.hlinkMouseOver = hlinkMouseOver
        self.rtl = rtl
        self.extLst = extLst


class TextTabStop(Serialisable):

    pos = unknown defintion for ST_Coordinate32
    algn = Set(values=(['l', 'ctr', 'r', 'dec']))

    def __init__(self,
                 pos=None,
                 algn=None,
                ):
        self.pos = pos
        self.algn = algn


class TextTabStopList(Serialisable):

    tab = Typed(expected_type=TextTabStop, allow_none=True)

    def __init__(self,
                 tab=None,
                ):
        self.tab = tab


class TextSpacing(Serialisable):

    pass

class TextParagraphProperties(Serialisable):

    marL = Coordinate()
    marR = Coordinate()
    lvl = Integer()
    indent = Coordinate()
    algn = Set(values=(['l', 'ctr', 'r', 'just', 'justLow', 'dist', 'thaiDist']))
    defTabSz = unknown defintion for ST_Coordinate32
    rtl = Bool(allow_none=True)
    eaLnBrk = Bool(allow_none=True)
    fontAlgn = Set(values=(['auto', 't', 'ctr', 'base', 'b']))
    latinLnBrk = Bool(allow_none=True)
    hangingPunct = Bool(allow_none=True)
    lnSpc = Typed(expected_type=TextSpacing, allow_none=True)
    spcBef = Typed(expected_type=TextSpacing, allow_none=True)
    spcAft = Typed(expected_type=TextSpacing, allow_none=True)
    tabLst = Typed(expected_type=TextTabStopList, allow_none=True)
    defRPr = Typed(expected_type=TextCharacterProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 marL=None,
                 marR=None,
                 lvl=None,
                 indent=None,
                 algn=None,
                 defTabSz=None,
                 rtl=None,
                 eaLnBrk=None,
                 fontAlgn=None,
                 latinLnBrk=None,
                 hangingPunct=None,
                 lnSpc=None,
                 spcBef=None,
                 spcAft=None,
                 tabLst=None,
                 defRPr=None,
                 extLst=None,
                ):
        self.marL = marL
        self.marR = marR
        self.lvl = lvl
        self.indent = indent
        self.algn = algn
        self.defTabSz = defTabSz
        self.rtl = rtl
        self.eaLnBrk = eaLnBrk
        self.fontAlgn = fontAlgn
        self.latinLnBrk = latinLnBrk
        self.hangingPunct = hangingPunct
        self.lnSpc = lnSpc
        self.spcBef = spcBef
        self.spcAft = spcAft
        self.tabLst = tabLst
        self.defRPr = defRPr
        self.extLst = extLst


class TextListStyle(Serialisable):

    defPPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl1pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl2pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl3pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl4pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl5pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl6pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl7pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl8pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl9pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 defPPr=None,
                 lvl1pPr=None,
                 lvl2pPr=None,
                 lvl3pPr=None,
                 lvl4pPr=None,
                 lvl5pPr=None,
                 lvl6pPr=None,
                 lvl7pPr=None,
                 lvl8pPr=None,
                 lvl9pPr=None,
                 extLst=None,
                ):
        self.defPPr = defPPr
        self.lvl1pPr = lvl1pPr
        self.lvl2pPr = lvl2pPr
        self.lvl3pPr = lvl3pPr
        self.lvl4pPr = lvl4pPr
        self.lvl5pPr = lvl5pPr
        self.lvl6pPr = lvl6pPr
        self.lvl7pPr = lvl7pPr
        self.lvl8pPr = lvl8pPr
        self.lvl9pPr = lvl9pPr
        self.extLst = extLst


class GeomGuide(Serialisable):

    name = String()
    fmla = String()

    def __init__(self,
                 name=None,
                 fmla=None,
                ):
        self.name = name
        self.fmla = fmla


class GeomGuideList(Serialisable):

    gd = Typed(expected_type=GeomGuide, allow_none=True)

    def __init__(self,
                 gd=None,
                ):
        self.gd = gd


class PresetTextShape(Serialisable):

    prst = Set(values=(['textNoShape', 'textPlain', 'textStop', 'textTriangle', 'textTriangleInverted', 'textChevron', 'textChevronInverted', 'textRingInside', 'textRingOutside', 'textArchUp', 'textArchDown', 'textCircle', 'textButton', 'textArchUpPour', 'textArchDownPour', 'textCirclePour', 'textButtonPour', 'textCurveUp', 'textCurveDown', 'textCanUp', 'textCanDown', 'textWave1', 'textWave2', 'textDoubleWave1', 'textWave4', 'textInflate', 'textDeflate', 'textInflateBottom', 'textDeflateBottom', 'textInflateTop', 'textDeflateTop', 'textDeflateInflate', 'textDeflateInflateDeflate', 'textFadeRight', 'textFadeLeft', 'textFadeUp', 'textFadeDown', 'textSlantUp', 'textSlantDown', 'textCascadeUp', 'textCascadeDown']))
    avLst = Typed(expected_type=GeomGuideList, allow_none=True)

    def __init__(self,
                 prst=None,
                 avLst=None,
                ):
        self.prst = prst
        self.avLst = avLst


class TextBodyProperties(Serialisable):

    rot = Integer()
    spcFirstLastPara = Bool(allow_none=True)
    vertOverflow = Set(values=(['overflow', 'ellipsis', 'clip']))
    horzOverflow = Set(values=(['overflow', 'clip']))
    vert = Set(values=(['horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl']))
    wrap = NoneSet(values=(['square']))
    lIns = unknown defintion for ST_Coordinate32
    tIns = unknown defintion for ST_Coordinate32
    rIns = unknown defintion for ST_Coordinate32
    bIns = unknown defintion for ST_Coordinate32
    numCol = Integer()
    spcCol = Coordinate()
    rtlCol = Bool(allow_none=True)
    fromWordArt = Bool(allow_none=True)
    anchor = Set(values=(['t', 'ctr', 'b', 'just', 'dist']))
    anchorCtr = Bool(allow_none=True)
    forceAA = Bool(allow_none=True)
    upright = Bool(allow_none=True)
    compatLnSpc = Bool(allow_none=True)
    prstTxWarp = Typed(expected_type=PresetTextShape, allow_none=True)
    scene3d = Typed(expected_type=Scene3D, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 rot=None,
                 spcFirstLastPara=None,
                 vertOverflow=None,
                 horzOverflow=None,
                 vert=None,
                 wrap=None,
                 lIns=None,
                 tIns=None,
                 rIns=None,
                 bIns=None,
                 numCol=None,
                 spcCol=None,
                 rtlCol=None,
                 fromWordArt=None,
                 anchor=None,
                 anchorCtr=None,
                 forceAA=None,
                 upright=None,
                 compatLnSpc=None,
                 prstTxWarp=None,
                 scene3d=None,
                 extLst=None,
                ):
        self.rot = rot
        self.spcFirstLastPara = spcFirstLastPara
        self.vertOverflow = vertOverflow
        self.horzOverflow = horzOverflow
        self.vert = vert
        self.wrap = wrap
        self.lIns = lIns
        self.tIns = tIns
        self.rIns = rIns
        self.bIns = bIns
        self.numCol = numCol
        self.spcCol = spcCol
        self.rtlCol = rtlCol
        self.fromWordArt = fromWordArt
        self.anchor = anchor
        self.anchorCtr = anchorCtr
        self.forceAA = forceAA
        self.upright = upright
        self.compatLnSpc = compatLnSpc
        self.prstTxWarp = prstTxWarp
        self.scene3d = scene3d
        self.extLst = extLst


class TextBody(Serialisable):

    bodyPr = Typed(expected_type=TextBodyProperties, )
    lstStyle = Typed(expected_type=TextListStyle, allow_none=True)
    p = Typed(expected_type=TextParagraph, )

    def __init__(self,
                 bodyPr=None,
                 lstStyle=None,
                 p=None,
                ):
        self.bodyPr = bodyPr
        self.lstStyle = lstStyle
        self.p = p


class NumFmt(Serialisable):

    formatCode = String()
    sourceLinked = Bool()

    def __init__(self,
                 formatCode=None,
                 sourceLinked=None,
                ):
        self.formatCode = formatCode
        self.sourceLinked = sourceLinked


class Tx(Serialisable):

    pass

class LayoutMode(Serialisable):

    val = Set(values=(['edge', 'factor']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class LayoutTarget(Serialisable):

    val = Set(values=(['inner', 'outer']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ManualLayout(Serialisable):

    layoutTarget = Typed(expected_type=LayoutTarget, allow_none=True)
    xMode = Typed(expected_type=LayoutMode, allow_none=True)
    yMode = Typed(expected_type=LayoutMode, allow_none=True)
    wMode = Typed(expected_type=LayoutMode, allow_none=True)
    hMode = Typed(expected_type=LayoutMode, allow_none=True)
    x = Float(nested=True, allow_none=True)
    y = Float(nested=True, allow_none=True)
    w = Float(nested=True, allow_none=True)
    h = Float(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 layoutTarget=None,
                 xMode=None,
                 yMode=None,
                 wMode=None,
                 hMode=None,
                 x=None,
                 y=None,
                 w=None,
                 h=None,
                 extLst=None,
                ):
        self.layoutTarget = layoutTarget
        self.xMode = xMode
        self.yMode = yMode
        self.wMode = wMode
        self.hMode = hMode
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.extLst = extLst


class Layout(Serialisable):

    manualLayout = Typed(expected_type=ManualLayout, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 manualLayout=None,
                 extLst=None,
                ):
        self.manualLayout = manualLayout
        self.extLst = extLst


class TrendlineLbl(Serialisable):

    layout = Typed(expected_type=Layout, allow_none=True)
    tx = Typed(expected_type=Tx, allow_none=True)
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 layout=None,
                 tx=None,
                 numFmt=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.layout = layout
        self.tx = tx
        self.numFmt = numFmt
        self.spPr = spPr
        self.txPr = txPr
        self.extLst = extLst


class Double(Serialisable):

    val = Float()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Period(Serialisable):

    val = Integer()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Order(Serialisable):

    val = Integer()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class TrendlineType(Serialisable):

    val = Set(values=(['exp', 'linear', 'log', 'movingAvg', 'poly', 'power']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Trendline(Serialisable):

    name = Typed(expected_type=String, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    trendlineType = Typed(expected_type=TrendlineType, )
    order = Typed(expected_type=Order, allow_none=True)
    period = Typed(expected_type=Period, allow_none=True)
    forward = Float(nested=True, allow_none=True)
    backward = Float(nested=True, allow_none=True)
    intercept = Float(nested=True, allow_none=True)
    dispRSqr = Bool(nested=True, allow_none=True)
    dispEq = Bool(nested=True, allow_none=True)
    trendlineLbl = Typed(expected_type=TrendlineLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 name=None,
                 spPr=None,
                 trendlineType=None,
                 order=None,
                 period=None,
                 forward=None,
                 backward=None,
                 intercept=None,
                 dispRSqr=None,
                 dispEq=None,
                 trendlineLbl=None,
                 extLst=None,
                ):
        self.name = name
        self.spPr = spPr
        self.trendlineType = trendlineType
        self.order = order
        self.period = period
        self.forward = forward
        self.backward = backward
        self.intercept = intercept
        self.dispRSqr = dispRSqr
        self.dispEq = dispEq
        self.trendlineLbl = trendlineLbl
        self.extLst = extLst


class DLbl(Serialisable):

    idx = Typed(expected_type=UnsignedInt, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 idx=None,
                 extLst=None,
                ):
        self.idx = idx
        self.extLst = extLst


class DLbls(Serialisable):

    dLbl = Typed(expected_type=DLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 dLbl=None,
                 extLst=None,
                ):
        self.dLbl = dLbl
        self.extLst = extLst


class PictureStackUnit(Serialisable):

    val = Float()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class PictureFormat(Serialisable):

    val = Set(values=(['stretch', 'stack', 'stackScale']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class PictureOptions(Serialisable):

    applyToFront = Bool(nested=True, allow_none=True)
    applyToSides = Bool(nested=True, allow_none=True)
    applyToEnd = Bool(nested=True, allow_none=True)
    pictureFormat = Typed(expected_type=PictureFormat, allow_none=True)
    pictureStackUnit = Typed(expected_type=PictureStackUnit, allow_none=True)

    def __init__(self,
                 applyToFront=None,
                 applyToSides=None,
                 applyToEnd=None,
                 pictureFormat=None,
                 pictureStackUnit=None,
                ):
        self.applyToFront = applyToFront
        self.applyToSides = applyToSides
        self.applyToEnd = applyToEnd
        self.pictureFormat = pictureFormat
        self.pictureStackUnit = pictureStackUnit


class Boolean(Serialisable):

    val = Bool(allow_none=True)

    def __init__(self,
                 val=None,
                ):
        self.val = val


class UnsignedInt(Serialisable):

    val = Integer()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class DPt(Serialisable):

    idx = Typed(expected_type=UnsignedInt, )
    invertIfNegative = Bool(nested=True, allow_none=True)
    marker = Typed(expected_type=Marker, allow_none=True)
    bubble3D = Bool(nested=True, allow_none=True)
    explosion = Typed(expected_type=UnsignedInt, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 idx=None,
                 invertIfNegative=None,
                 marker=None,
                 bubble3D=None,
                 explosion=None,
                 spPr=None,
                 pictureOptions=None,
                 extLst=None,
                ):
        self.idx = idx
        self.invertIfNegative = invertIfNegative
        self.marker = marker
        self.bubble3D = bubble3D
        self.explosion = explosion
        self.spPr = spPr
        self.pictureOptions = pictureOptions
        self.extLst = extLst


class Extension(Serialisable):

    uri = String()

    def __init__(self,
                 uri=None,
                ):
        self.uri = uri


class ExtensionList(Serialisable):

    ext = Typed(expected_type=Extension, allow_none=True)

    def __init__(self,
                 ext=None,
                ):
        self.ext = ext


class Color(Serialisable):

    pass

class Bevel(Serialisable):

    w = Float()
    h = Float()
    prst = Set(values=(['relaxedInset', 'circle', 'slope', 'cross', 'angle', 'softRound', 'convex', 'coolSlant', 'divot', 'riblet', 'hardEdge', 'artDeco']))

    def __init__(self,
                 w=None,
                 h=None,
                 prst=None,
                ):
        self.w = w
        self.h = h
        self.prst = prst


class Shape3D(Serialisable):

    z = unknown defintion for ST_Coordinate
    extrusionH = Float()
    contourW = Float()
    prstMaterial = Set(values=(['legacyMatte', 'legacyPlastic', 'legacyMetal', 'legacyWireframe', 'matte', 'plastic', 'metal', 'warmMatte', 'translucentPowder', 'powder', 'dkEdge', 'softEdge', 'clear', 'flat', 'softmetal']))
    bevelT = Typed(expected_type=Bevel, allow_none=True)
    bevelB = Typed(expected_type=Bevel, allow_none=True)
    extrusionClr = Typed(expected_type=Color, allow_none=True)
    contourClr = Typed(expected_type=Color, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 z=None,
                 extrusionH=None,
                 contourW=None,
                 prstMaterial=None,
                 bevelT=None,
                 bevelB=None,
                 extrusionClr=None,
                 contourClr=None,
                 extLst=None,
                ):
        self.z = z
        self.extrusionH = extrusionH
        self.contourW = contourW
        self.prstMaterial = prstMaterial
        self.bevelT = bevelT
        self.bevelB = bevelB
        self.extrusionClr = extrusionClr
        self.contourClr = contourClr
        self.extLst = extLst


class Vector3D(Serialisable):

    dx = unknown defintion for ST_Coordinate
    dy = unknown defintion for ST_Coordinate
    dz = unknown defintion for ST_Coordinate

    def __init__(self,
                 dx=None,
                 dy=None,
                 dz=None,
                ):
        self.dx = dx
        self.dy = dy
        self.dz = dz


class Point3D(Serialisable):

    x = unknown defintion for ST_Coordinate
    y = unknown defintion for ST_Coordinate
    z = unknown defintion for ST_Coordinate

    def __init__(self,
                 x=None,
                 y=None,
                 z=None,
                ):
        self.x = x
        self.y = y
        self.z = z


class Backdrop(Serialisable):

    anchor = Typed(expected_type=Point3D, )
    norm = Typed(expected_type=Vector3D, )
    up = Typed(expected_type=Vector3D, )
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 anchor=None,
                 norm=None,
                 up=None,
                 extLst=None,
                ):
        self.anchor = anchor
        self.norm = norm
        self.up = up
        self.extLst = extLst


class LightRig(Serialisable):

    rig = Set(values=(['legacyFlat1', 'legacyFlat2', 'legacyFlat3', 'legacyFlat4', 'legacyNormal1', 'legacyNormal2', 'legacyNormal3', 'legacyNormal4', 'legacyHarsh1', 'legacyHarsh2', 'legacyHarsh3', 'legacyHarsh4', 'threePt', 'balanced', 'soft', 'harsh', 'flood', 'contrasting', 'morning', 'sunrise', 'sunset', 'chilly', 'freezing', 'flat', 'twoPt', 'glow', 'brightRoom']))
    dir = Set(values=(['tl', 't', 'tr', 'l', 'r', 'bl', 'b', 'br']))
    rot = Typed(expected_type=SphereCoords, allow_none=True)

    def __init__(self,
                 rig=None,
                 dir=None,
                 rot=None,
                ):
        self.rig = rig
        self.dir = dir
        self.rot = rot


class SphereCoords(Serialisable):

    lat = Integer()
    lon = Integer()
    rev = Integer()

    def __init__(self,
                 lat=None,
                 lon=None,
                 rev=None,
                ):
        self.lat = lat
        self.lon = lon
        self.rev = rev


class Camera(Serialisable):

    prst = Set(values=(['legacyObliqueTopLeft', 'legacyObliqueTop', 'legacyObliqueTopRight', 'legacyObliqueLeft', 'legacyObliqueFront', 'legacyObliqueRight', 'legacyObliqueBottomLeft', 'legacyObliqueBottom', 'legacyObliqueBottomRight', 'legacyPerspectiveTopLeft', 'legacyPerspectiveTop', 'legacyPerspectiveTopRight', 'legacyPerspectiveLeft', 'legacyPerspectiveFront', 'legacyPerspectiveRight', 'legacyPerspectiveBottomLeft', 'legacyPerspectiveBottom', 'legacyPerspectiveBottomRight', 'orthographicFront', 'isometricTopUp', 'isometricTopDown', 'isometricBottomUp', 'isometricBottomDown', 'isometricLeftUp', 'isometricLeftDown', 'isometricRightUp', 'isometricRightDown', 'isometricOffAxis1Left', 'isometricOffAxis1Right', 'isometricOffAxis1Top', 'isometricOffAxis2Left', 'isometricOffAxis2Right', 'isometricOffAxis2Top', 'isometricOffAxis3Left', 'isometricOffAxis3Right', 'isometricOffAxis3Bottom', 'isometricOffAxis4Left', 'isometricOffAxis4Right', 'isometricOffAxis4Bottom', 'obliqueTopLeft', 'obliqueTop', 'obliqueTopRight', 'obliqueLeft', 'obliqueRight', 'obliqueBottomLeft', 'obliqueBottom', 'obliqueBottomRight', 'perspectiveFront', 'perspectiveLeft', 'perspectiveRight', 'perspectiveAbove', 'perspectiveBelow', 'perspectiveAboveLeftFacing', 'perspectiveAboveRightFacing', 'perspectiveContrastingLeftFacing', 'perspectiveContrastingRightFacing', 'perspectiveHeroicLeftFacing', 'perspectiveHeroicRightFacing', 'perspectiveHeroicExtremeLeftFacing', 'perspectiveHeroicExtremeRightFacing', 'perspectiveRelaxed', 'perspectiveRelaxedModerately']))
    fov = Integer()
    zoom = unknown defintion for ST_PositivePercentage
    rot = Typed(expected_type=SphereCoords, allow_none=True)

    def __init__(self,
                 prst=None,
                 fov=None,
                 zoom=None,
                 rot=None,
                ):
        self.prst = prst
        self.fov = fov
        self.zoom = zoom
        self.rot = rot


class Scene3D(Serialisable):

    camera = Typed(expected_type=Camera, )
    lightRig = Typed(expected_type=LightRig, )
    backdrop = Typed(expected_type=Backdrop, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 camera=None,
                 lightRig=None,
                 backdrop=None,
                 extLst=None,
                ):
        self.camera = camera
        self.lightRig = lightRig
        self.backdrop = backdrop
        self.extLst = extLst


class OfficeArtExtensionList(Serialisable):

    pass

class LineEndProperties(Serialisable):

    type = NoneSet(values=(['triangle', 'stealth', 'diamond', 'oval', 'arrow']))
    w = Set(values=(['sm', 'med', 'lg']))
    len = Set(values=(['sm', 'med', 'lg']))

    def __init__(self,
                 type=None,
                 w=None,
                 len=None,
                ):
        self.type = type
        self.w = w
        self.len = len


class LineProperties(Serialisable):

    w = Coordinate()
    cap = Set(values=(['rnd', 'sq', 'flat']))
    cmpd = Set(values=(['sng', 'dbl', 'thickThin', 'thinThick', 'tri']))
    algn = Set(values=(['ctr', 'in']))
    headEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    tailEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 w=None,
                 cap=None,
                 cmpd=None,
                 algn=None,
                 headEnd=None,
                 tailEnd=None,
                 extLst=None,
                ):
        self.w = w
        self.cap = cap
        self.cmpd = cmpd
        self.algn = algn
        self.headEnd = headEnd
        self.tailEnd = tailEnd
        self.extLst = extLst


class PositiveSize2D(Serialisable):

    cx = Float()
    cy = Float()

    def __init__(self,
                 cx=None,
                 cy=None,
                ):
        self.cx = cx
        self.cy = cy


class Point2D(Serialisable):

    x = unknown defintion for ST_Coordinate
    y = unknown defintion for ST_Coordinate

    def __init__(self,
                 x=None,
                 y=None,
                ):
        self.x = x
        self.y = y


class Transform2D(Serialisable):

    rot = Integer()
    flipH = Bool(allow_none=True)
    flipV = Bool(allow_none=True)
    off = Typed(expected_type=Point2D, allow_none=True)
    ext = Typed(expected_type=PositiveSize2D, allow_none=True)

    def __init__(self,
                 rot=None,
                 flipH=None,
                 flipV=None,
                 off=None,
                 ext=None,
                ):
        self.rot = rot
        self.flipH = flipH
        self.flipV = flipV
        self.off = off
        self.ext = ext


class ShapeProperties(Serialisable):

    bwMode = Set(values=(['clr', 'auto', 'gray', 'ltGray', 'invGray', 'grayWhite', 'blackGray', 'blackWhite', 'black', 'white', 'hidden']))
    xfrm = Typed(expected_type=Transform2D, allow_none=True)
    ln = Typed(expected_type=LineProperties, allow_none=True)
    scene3d = Typed(expected_type=Scene3D, allow_none=True)
    sp3d = Typed(expected_type=Shape3D, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 bwMode=None,
                 xfrm=None,
                 ln=None,
                 scene3d=None,
                 sp3d=None,
                 extLst=None,
                ):
        self.bwMode = bwMode
        self.xfrm = xfrm
        self.ln = ln
        self.scene3d = scene3d
        self.sp3d = sp3d
        self.extLst = extLst


class MarkerSize(Serialisable):

    val = Integer()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class MarkerStyle(Serialisable):

    val = NoneSet(values=(['circle', 'dash', 'diamond', 'dot', 'picture', 'plus', 'square', 'star', 'triangle', 'x', 'auto']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Marker(Serialisable):

    symbol = Typed(expected_type=MarkerStyle, allow_none=True)
    size = Typed(expected_type=MarkerSize, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 symbol=None,
                 size=None,
                 spPr=None,
                 extLst=None,
                ):
        self.symbol = symbol
        self.size = size
        self.spPr = spPr
        self.extLst = extLst


class ScatterSer(Serialisable):

    marker = Typed(expected_type=Marker, allow_none=True)
    dPt = Typed(expected_type=DPt, allow_none=True)
    dLbls = Typed(expected_type=DLbls, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrBars, allow_none=True)
    xVal = Typed(expected_type=AxDataSource, allow_none=True)
    yVal = Typed(expected_type=NumDataSource, allow_none=True)
    smooth = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 marker=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 xVal=None,
                 yVal=None,
                 smooth=None,
                 extLst=None,
                ):
        self.marker = marker
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.xVal = xVal
        self.yVal = yVal
        self.smooth = smooth
        self.extLst = extLst

