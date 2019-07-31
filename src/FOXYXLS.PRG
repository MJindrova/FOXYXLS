* FOXYXLS Class
* Creates pure XLS files without the need of having MS-OFFICE installed
* Author: VFPIMAGING
* Based on Serhiy Perevoznyk's work published in his blog
* http://delphi32.blogspot.com.br/2011/06/generate-excel-files-without-using.html

#INCLUDE foxyxls.h

#DEFINE VFP_VERSION VAL(SUBS(VERSION(),LEN("Visual FoxPro ")+1,2))
#DEFINE ARRAY_LIMIT 65000

#IF VFP_VERSION<7
 PROCEDURE FOXYXLS_APINIT
   DECLARE memcpy IN msvcr71.DLL AS FOXYXLS_memcpy STRING@ DEST, DOUBLE@ SOURCE, INTEGER LENGTH
 ENDPROC 
#ENDIF

#IF VFP_VERSION>=7
 PROCEDURE FOXYXLS_memcpy
    LPARAMETERS m.lcString, m.lnSource, m.liLen
    DECLARE memcpy IN msvcr71.DLL AS FOXYXLS_memcpy STRING@ DEST, DOUBLE@ SOURCE, INTEGER LENGTH
    =FOXYXLS_memcpy(@m.lcString, m.lnSource, m.liLen)
 ENDPROC   
#ENDIF

#IF VFP_VERSION<9
#DEFINE BINTOC BINTOC_LT9

FUNCTION BINTOC_LT9
   LPARAMETERS m.pnValue, m.pcFlags
   LOCAL m.lcResult
   m.lcResult = ""
   DO CASE
      CASE m.pcFlags = "4RS"
           m.lcResult = CHR(BITAND(m.pnValue, 0xFF))+CHR(BITRSHIFT(BITAND(m.pnValue, 0xFF00), 8))+;
                        CHR(BITRSHIFT(BITAND(m.pnValue, 0xFF0000), 16))+CHR(BITRSHIFT(BITAND(m.pnValue, 0xFF000000), 24))

      CASE m.pcFlags = "2RS"
           m.lcResult = CHR(BITAND(m.pnValue, 0xFF))+CHR(BITRSHIFT(BITAND(m.pnValue, 0xFF00), 8))

      CASE m.pcFlags = "8SB"
           * Credits for Rick Hodgin
           m.lcResult = SPACE(8)
           FOXYXLS_memcpy(@m.lcResult, @m.pnValue, 8)
           
   ENDCASE
   RETURN m.lcResult
ENDFUNC
#ENDIF

DEFINE CLASS FoxyXLS AS CUSTOM
   cFile     = ""
   cDateType = "B"
   cAuthor   = ""
   nCodePage           = 1252
   _nCurrFXRecord      = 20
   cDefaultFontName    = "Arial"
   nDefaultFontSize    = 10
   nDefaultColumnWidth = 0
   nDefaultRowHeight   = 0
   cAFXStyles=""                            && Cell styles
   cACells=""                               && Cell metadata data, values are in aCells() array
   _iAGroup=1                               && Cell values group array index (VFP6-8)  
   _iAIndex=0                               && Cell value index
   _cARRAY_LIMIT=LTRIM(STR(ARRAY_LIMIT,11)) && Cell Array limit (65000)
   _hFile=0
   
   
   DIMENSION aFonts[1,1], aCells[ARRAY_LIMIT], aFmts[1,1], aColWidths[1,1], aColors[1,1]

   FUNCTION Init
      *
      * FoxyXLS::Init
      *
      #IF VFP_VERSION<7
       =FOXYXLS_APINIT()
      #ENDIF

      This.aFonts[1,1]=.F.
      STORE "" TO This.aCells
      This.aFmts[1,1]=""
      This.aColWidths[1,1]=""
      This.aColors[1,1]=""

      This.PrepareFormatArray()
      This.PrepareColorsArray()
      This.PrepareFontArray()

      LOCAL m.lcSetDate, m.lcDateType, m.liSele
      m.lcSetDate = UPPER(ALLTRIM(SET("Date")))

      DO CASE
         CASE INLIST(m.lcSetDate, "BRITISH", "FRENCH", "DMY")
              m.lcDateType = "B"

         CASE INLIST(m.lcSetDate, "AMERICAN", "USA", "MDY")
              m.lcDateType = "A"

         OTHERWISE
              m.lcDateType = "B"
      ENDCASE
      This.cDateType = m.lcDateType


      m.liSele=SELECT()
      This.cAFXStyles="FXLS_FXS_"+SYS(2015)
      CREATE CURSOR (This.cAFXStyles) (FXSTYLE I, FONTINDEX I, FORMATINDEX I, ALIGN I, BACKCOLOR I)
      SELECT (This.cAFXStyles)
      INDEX ON STR(FONTINDEX,5)+STR(FORMATINDEX,5)+STR(ALIGN,3)+STR(BACKCOLOR,5) TAG "I01"

      This.cACells="FXLS_CELSS_"+SYS(2015)
      CREATE CURSOR (This.cACells) (ROWNO I, COLNO I, FXSTYLE I, AGROUP I, AINDEX I)

      SELECT (m.liSele)
   ENDFUNC


   FUNCTION Destroy
      *
      * FoxyXLS::Destroy
      *
      USE IN (This.cAFXStyles)
      USE IN (This.cACells)
   ENDFUNC


   FUNCTION Clear
      *
      * FoxyXLS::Clear
      *
      LOCAL m.lcSAFETY
      m.lcSAFETY=SET("SAFETY")
      SELECT (This.cAFXStyles)
      ZAP
      SELECT (This.cACells)
      ZAP
      SET SAFETY &lcSAFETY.
      This._iAGroup=1
      This._iAIndex=0
      This.PrepareFontArray()
      This.PrepareFormatArray()
   ENDFUNC


   FUNCTION PrepareColorsArray
      *
      * FoxyXLS::PrepareColorsArray()
      *
      DIMENSION This.aColors(17,3)
      WITH This
      .aColors(1,1) = "WHITE" && 0xFFFFFF
      .aColors(1,2) = 0x01    && ForeColor
      .aColors(1,3) = 0xC041  && BackColor

      .aColors(2,1) = "RED"   && 0xFF0000  
      .aColors(2,2) = 0x02    && ForeColor
      .aColors(2,3) = 0xC081  && BackColor

      .aColors(3,1) = "GREEN" && 0x00FF00
      .aColors(3,2) = 0x03    && ForeColor
      .aColors(3,3) = 0xC0C1  && BackColor

      .aColors(4,1) = "BLUE"  && 0x0000FF
      .aColors(4,2) = 0x04    && ForeColor
      .aColors(4,3) = 0xC101  && BackColor

      .aColors(5,1) = "YELLOW"&& 0xFFFF00
      .aColors(5,2) = 0x05    && ForeColor
      .aColors(5,3) = 0xC141  && BackColor

      .aColors(6,1) = "MAGENTA" && 0xFF00FF
      .aColors(6,2) = 0x06    && ForeColor
      .aColors(6,3) = 0xC181  && BackColor

      .aColors(7,1) = "CYAN"  && 0x00FFFF
      .aColors(7,2) = 0x07    && ForeColor
      .aColors(7,3) = 0xC1C1  && BackColor

      .aColors(8,1) = "DARKRED" && 0x800000
      .aColors(8,2) = 0x10    && ForeColor
      .aColors(8,3) = 0xC401  && BackColor

      .aColors(9,1) = "DARKGREEN" && 0x008000 
      .aColors(9,2) = 0x11    && ForeColor
      .aColors(9,3) = 0xC441  && BackColor

      .aColors(10,1) = "DARKBLUE" && 0x000080
      .aColors(10,2) = 0x12    && ForeColor
      .aColors(10,3) = 0xC481  && BackColor

      .aColors(11,1) = "OLIVE" && 0x808000
      .aColors(11,2) = 0x13    && ForeColor
      .aColors(11,3) = 0xC4C1  && BackColor

      .aColors(12,1) = "PURPLE" && 0x800080
      .aColors(12,2) = 0x14    && ForeColor
      .aColors(12,3) = 0xC501  && BackColor

      .aColors(13,1) = "TEAL"  && 0x008080
      .aColors(13,2) = 0x15    && ForeColor
      .aColors(13,3) = 0xC541  && BackColor

      .aColors(14,1) = "SILVER" && 0xC0C0C0
      .aColors(14,2) = 0x16    && ForeColor
      .aColors(14,3) = 0xC581  && BackColor

      .aColors(15,1) = "GRAY"  && 0x808080
      .aColors(15,2) = 0x17    && ForeColor
      .aColors(15,3) = 0xC5C1  && BackColor

      .aColors(16,1) = "BLACK" && 0x000000
      .aColors(16,2) = 0x00    && ForeColor
      .aColors(16,3) = 0xC001  && BackColor

      .aColors(17,1) = "AUTOMATIC" && NULL
      .aColors(17,2) = 0x7FFF  && ForeColor
      .aColors(17,3) = 0xCE00  && BackColor
      ENDWITH 
   ENDFUNC 


   FUNCTION PrepareFormatArray
      *
      * FoxyXLS::PrepareFormatArray()
      *
      * Prepare the Formatting record
      DIMENSION This.aFmts(38)
      WITH This
      .aFmts(1)=("General")
      .aFmts(2)=("0")
      .aFmts(3)=("0.00")
      .aFmts(4)=("#,##0")
      .aFmts(5)=("#,##0.00")
      .aFmts(6)=("($#,##0_);($#,##0)")
      .aFmts(7)=("($#,##0_);[Red]($#,##0)")
      .aFmts(8)=("($#,##0.00_);($#,##0.00)")
      .aFmts(9)=("($#,##0.00_);[Red]($#,##0.00)")
      .aFmts(10)=("0%")
      .aFmts(11)=("0.00%")
      .aFmts(12)=("0.00E+00")
      .aFmts(13)=("# ?/?")
      .aFmts(14)=("# ??/??")
      .aFmts(15)=("m/d/yy")
      .aFmts(16)=("d-mmm-yy")
      .aFmts(17)=("d-mmm")
      .aFmts(18)=("mmm-yy")
      .aFmts(19)=("h:mm AM/PM")
      .aFmts(20)=("h:mm:ss AM/PM")
      .aFmts(21)=("h:mm")
      .aFmts(22)=("h:mm:ss")
      .aFmts(23)=("m/d/yy h:mm")
      .aFmts(24)=("(#,##0_);(#,##0)")
      .aFmts(25)=("(#,##0_);[Red](#,##0)")
      .aFmts(26)=("(#,##0.00_);(#,##0.00)")
      .aFmts(27)=("(#,##0.00_);[Red](#,##0.00)")
      .aFmts(28)=('_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)')
      .aFmts(29)=('_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)')
      .aFmts(30)=('_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)')
      .aFmts(31)=('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)')
      .aFmts(32)=("mm:ss")
      .aFmts(33)=("[h]:mm:ss")
      .aFmts(34)=("mm:ss.0")
      .aFmts(35)=("##0.0E+0")
      .aFmts(36)=("@")
      .aFmts(37)=("dd/mm/yyyy")
      .aFmts(38)=("$#,##0.00")
      ENDWITH 
   ENDFUNC 


   FUNCTION PrepareFontArray
      *
      * FoxyXLS::PrepareFontArray()
      *
      * Prepare the fonts Array
      LOCAL m.lnXLSColor
      This.AddFont(This.GetFontFullName("Arial,10,  ,BLACK", , @m.lnXLSColor), m.lnXLSColor)  && MJ 2018-11-01
      This.AddFont(This.GetFontFullName("Arial,10,B ,BLACK", , @m.lnXLSColor), m.lnXLSColor)  && MJ 2018-11-01
      This.AddFont(This.GetFontFullName("Arial,10,I ,BLACK", , @m.lnXLSColor), m.lnXLSColor)  && MJ 2018-11-01
      This.AddFont(This.GetFontFullName("Arial,10,BI,BLACK", , @m.lnXLSColor), m.lnXLSColor)  && MJ 2018-11-01
   ENDFUNC 


   FUNCTION AddFont
      *
      * FoxyXLS::AddFont()
      *
      LPARAMETERS m.tcFontFull, m.tnXLSColor
      
      LOCAL m.liCNT, m.lcFont, m.lnSize, m.lcStyle, m.lnFontIndex, m.lcColor, m.lii
      LOCAL m.laFontA(1)

      #IF VFP_VERSION>=7
       m.lnFontIndex=ASCAN(This.aFonts, m.tcFontFull, 1, -1, 1, 4+2+1+8)
      #ENDIF

      #IF VFP_VERSION<7
       m.lnFontIndex=0
       FOR m.lii=1 TO ALEN(This.aFonts, 1)
           IF NOT EMPTY(This.aFonts(m.lii, 1)) AND UPPER(This.aFonts(m.lii, 1))==UPPER(m.tcFontFull)
              m.lnFontIndex=m.lii
              EXIT
           ENDIF
       NEXT
      #ENDIF

      IF m.lnFontIndex = 0
         m.liCNT=ALINES(m.laFontA, CHRTRAN(m.tcFontFull, ",", CHR(10)))
         m.lcFont  = ALLTRIM(m.laFontA(1))
         m.lnSize  = IIF(m.liCNT>=2, VAL(ALLTRIM(m.laFontA(2))), 10)
         m.lcStyle = IIF(m.liCNT>=3, UPPER(ALLTRIM(m.laFontA(3))), "")
         m.lcColor = IIF(m.liCNT>=4, UPPER(ALLTRIM(m.laFontA(4))), "")

         m.lnFontTotal = IIF(TYPE("This.aFonts(1)")="L", 0, ALEN(This.aFonts,1))
         m.lnFontIndex = m.lnFontTotal + 1
         DIMENSION This.aFonts(m.lnFontIndex, 6)
         This.aFonts(m.lnFontIndex, 1) = m.tcFontFull
         This.aFonts(m.lnFontIndex, 2) = m.lcFont
         This.aFonts(m.lnFontIndex, 3) = m.lnSize
         This.aFonts(m.lnFontIndex, 4) = m.lcStyle
         This.aFonts(m.lnFontIndex, 5) = m.lcColor
         This.aFonts(m.lnFontIndex, 6) = m.tnXLSColor
      ENDIF
      RETURN IIF(m.lnFontIndex<=4, MAX(m.lnFontIndex-1,0), m.lnFontIndex)
   ENDFUNC


   FUNCTION GetFont
      *
      * FoxyXLS::GetFont()
      *
      LPARAMETERS m.tcFontFull
      
      LOCAL m.liCNT, m.lcFont, m.lnSize, m.lcStyle, m.lnFontIndex, m.lnColor, m.lii

      #IF VFP_VERSION>=7
       m.lnFontIndex=ASCAN(This.aFonts, m.tcFontFull, 1, -1, 1, 4+2+1+8)
      #ENDIF

      #IF VFP_VERSION<7
       m.lnFontIndex=0
       FOR m.lii=1 TO ALEN(This.aFonts, 1)
           IF UPPER(This.aFonts(m.lii, 1))==UPPER(m.tcFontFull)
              m.lnFontIndex=m.lii
              EXIT
           ENDIF
       NEXT
      #ENDIF
      RETURN IIF(m.lnFontIndex<=4, m.lnFontIndex-1, m.lnFontIndex) && return -1 if font not found
   ENDFUNC


   PROCEDURE GetFontFullName
      *
      * FoxyXLS::GetFontFullName()
      *
      LPARAMETERS m.tcFontFull, m.tnForeColor, m.tnXLSColor

      * Check if only the forecolor was passed
      IF EMPTY(m.tcFontFull) AND m.tnForeColor > 0
         IF m.tnForeColor<=ALEN(This.aColors,1)
            m.tnXlsColor  = This.aColors(m.tnForeColor, 2)
            m.lcFontFull  = ",,," + This.aColors(m.tnForeColor, 1)
         ENDIF
         RETURN m.lcFontFull
      ENDIF 

      
      * Working with the supposition NOT EMPTY(tcFontFull)
       
      * Getting the correct Forecolor value
      * And preparing the internal tables to work with it
      LOCAL m.lcFontFull, m.lcColorName, m.lnColor, m.lcFont, m.lnSize, m.lcStyle, m.lcColor, m.lii
      LOCAL ARRAY m.laFontA(1)

      =ALINES(m.laFontA, CHRTRAN(STRTRAN(m.tcFontFull,',,',', ,'), ",", CHR(10)))
      m.lcFont  = ALLTRIM(m.laFontA(1))
      IF EMPTY(m.lcFont)
         m.lcFont = This.cDefaultFontName
      ENDIF 
      m.lnSize  = INT(VAL(ALLTRIM(m.laFontA(2))))
      IF EMPTY(m.lnSize)
         m.lnSize = This.nDefaultFontSize
      ENDIF 
      m.lcStyle = UPPER(ALLTRIM(m.laFontA(3)))
      m.lcColor = UPPER(ALLTRIM(m.laFontA(4)))

      m.tcFontFull = m.lcFont + "," + TRANSFORM(m.lnSize) + "," + IIF(EMPTY(m.lcStyle), " ",m.lcStyle) + "," && Must be space for VFP 6

      DO CASE
         CASE EMPTY(m.lcColor) AND m.tnForeColor = 0 && No color passed
              m.tnXlsColor  = 0
              m.lcFontFull  = m.tcFontFull

         CASE NOT EMPTY(m.lcColor)
              #IF VFP_VERSION>=7
               m.lnIndex = ASCAN(This.aColors, m.lcColor, 1, 0, 1, 4+2+1+8)
              #ENDIF

              #IF VFP_VERSION<7
               m.lnIndex=0
               FOR m.lii=1 TO ALEN(This.aColors, 1)
                   IF UPPER(This.aColors(m.lii, 1))==UPPER(m.lcColor)
                      m.lnIndex=m.lii
                      EXIT
                   ENDIF
               NEXT
              #ENDIF
         
              IF m.lnIndex = 0 && Not found, the ignore the color
                 m.tnXlsColor  = 0
                 m.lcFontFull  = m.tcFontFull
              ELSE 
                 && Found the color
                 m.tnXlsColor  = This.aColors(m.lnIndex, 2)
                 m.lcFontFull  = m.tcFontFull + m.lcColor
              ENDIF 

         OTHERWISE
              m.tnXlsColor  = 0
              m.lcFontFull  = m.tcFontFull
      ENDCASE

      RETURN m.lcFontFull
   ENDPROC 


   PROCEDURE AddCell
      *
      * FoxyXLS::AddCell()
      *
      LPARAMETERS m.tnRow, m.tnCol, m.tuValue, m.tcFontFull, m.tcFormat, m.tnAlign, m.tnForeColor, m.tnBackColor, m.lii
      
      IF m.tnRow>16384 OR m.tnCol>256
         RETURN
      ENDIF
      
      #IF VFP_VERSION<8
       m.tcFontFull  = IIF(EMPTY(m.tcFontFull), "", m.tcFontFull)
       m.tnForeColor = IIF(EMPTY(m.tnForeColor), 0, m.tnForeColor)
       m.tnAlign     = IIF(EMPTY(m.tnAlign), 0, m.tnAlign)
       m.tnBackColor = IIF(EMPTY(m.tnBackColor), 0, m.tnBackColor) 
      #ELSE
       m.tcFontFull  = EVL(m.tcFontFull, "")
       m.tnForeColor = EVL(m.tnForeColor, 0)
       m.tnAlign     = EVL(m.tnAlign, 0)
       m.tnBackColor = EVL(m.tnBackColor, 0) 
      #ENDIF
      * Prepare the Cells array, having the next record initialized to receive the information
      LOCAL m.lnTotCells, m.i, m.lcFontFull, m.lnXLSColor, m.lcValueType, m.lnFormatIndex, m.luFmtIndex
      
      m.lcFontFull  = ""
      m.lnXLSColor  = 0
      IF EMPTY(m.tcFontFull) AND EMPTY(m.tnForeColor)
      ELSE
         m.lcFontFull = This.GetFontFullName(m.tcFontFull, m.tnForeColor, @m.lnXLSColor)
      ENDIF 

      IF EMPTY(m.lcFontFull)
         m.lnFontIndex = 0
      ELSE 
         m.lnFontIndex = This.GetFont(m.lcFontFull)
         IF m.lnFontIndex = -1
            m.lnFontIndex = This.AddFont(m.lcFontFull, m.lnXLSColor)
         ENDIF
      ENDIF 

      m.lcValueType = TYPE("m.tuValue")
      IF m.lcValueType = "D" AND EMPTY(m.tcFormat)
         DO CASE
            CASE This.cDateType = "A"
                 m.tcFormat = "m/d/yy"
                 
            CASE This.cDateType = "B"
                 m.tcFormat = "dd/mm/yyyy"

            OTHERWISE
         ENDCASE
      ENDIF 
      IF m.lcValueType = "T" AND EMPTY(m.tcFormat)
         DO CASE
            CASE This.cDateType = "A"
                 m.tcFormat = "m/d/yy h:m:s"
                 
            CASE This.cDateType = "B"
                 m.tcFormat = "dd/mm/yyyy hh:mm:ss"

            OTHERWISE
         ENDCASE
      ENDIF 

      * Getting the correct Format Index
      m.luFmtIndex = m.tcFormat
      DO CASE
         CASE TYPE("m.luFmtIndex") = "N"
              m.lnFormatIndex = m.luFmtIndex

         CASE TYPE("m.luFmtIndex") = "C"
              #IF VFP_VERSION>=7
               m.lnFormatIndex = ASCAN(This.aFmts, ALLTRIM(m.luFmtIndex), 1, -1, 1, 4+2+1+8)
              #ENDIF

              #IF VFP_VERSION<7
               m.lnFormatIndex=0
               FOR m.lii=1 TO ALEN(This.aFmts, 1)
                   IF UPPER(This.aFmts(m.lii, 1))==UPPER(ALLTRIM(m.luFmtIndex))
                      m.lnFormatIndex=m.lii
                      EXIT
                   ENDIF
               NEXT
              #ENDIF
              IF m.lnFormatIndex=0
                 m.lnFormatIndex=ALEN(This.aFmts,1)+1
                 DIME This.aFmts(m.lnFormatIndex)
                 This.aFmts(m.lnFormatIndex)=m.luFmtIndex
              ENDIF
              m.lnFormatIndex = MAX(m.lnFormatIndex - 1, 0)
              
         OTHERWISE
              m.lnFormatIndex = 0
      ENDCASE

      
      IF NOT SEEK(STR(m.lnFontIndex, 5)+STR(m.lnFormatIndex, 5)+STR(m.tnAlign, 3)+STR(m.tnBackColor, 5), This.cAFXStyles)
         INSERT INTO (This.cAFXStyles) (FXSTYLE, FONTINDEX, FORMATINDEX, ALIGN, BACKCOLOR);
           VALUES (RECCOUNT(This.cAFXStyles)+1+This._nCurrFXRecord, m.lnFontIndex, m.lnFormatIndex, m.tnAlign, m.tnBackColor)
      ENDIF 

      #IF VFP_VERSION<9
       IF This._iAIndex=ARRAY_LIMIT
          This._iAGroup=This._iAGroup+1
          This.AddProperty("aCells"+LTRIM(STR(This._iAGroup,2))+"("+This._cARRAY_LIMIT+")")
          STORE 1 TO This._iAIndex, m.i
       ELSE
          STORE This._iAIndex+1 TO m.i, This._iAIndex
       ENDIF
      #ELSE
       IF This._iAIndex>0 AND This._iAIndex%ARRAY_LIMIT=0
          DIMENSION This.aCells(This._iAIndex+ARRAY_LIMIT)
       ENDIF
       STORE This._iAIndex+1 TO m.i, This._iAIndex
      #ENDIF

      INSERT INTO (This.cACells) (ROWNO, COLNO, FXSTYLE, AGROUP, AINDEX);
        VALUES (m.tnRow, m.tnCol, EVALUATE(This.cAFXStyles+".FXSTYLE"), This._iAGroup, m.i)

      #IF VFP_VERSION<9
       STORE m.tuValue TO (IIF(This._iAGroup=1, "This.aCells(m.i)", "This.aCells"+LTRIM(STR(This._iAGroup,2))+"(m.i)"))
      #ELSE
       This.aCells(m.i) = m.tuValue
      #ENDIF

   ENDPROC 


   FUNCTION SetColumnWidth
      *
      * FoxyXLS::SetColumnWidth()
      *
      LPARAMETERS m.tnColumn, m.tnWidth
      
      LOCAL m.lnCols, m.lnNextCol
      m.lnCols = ALEN(This.aColWidths, 1)
      IF TYPE("This.aColWidths(1)") <> "N"
         m.lnNextCol = 1
      ELSE 
         m.lnNextCol = m.lnCols + 1
      ENDIF
      DIMENSION This.aColWidths(m.lnNextCol, 2)
      This.aColWidths(m.lnNextCol, 1) = m.tnColumn
      This.aColWidths(m.lnNextCol, 2) = m.tnWidth
   ENDFUNC 


   FUNCTION WriteFile
      *
      * FoxyXLS::WriteFile()
      *
      LPARAMETERS m.tcFile
      
      IF EMPTY(m.tcFile)
         m.tcFile = This.cFile
      ENDIF 
      LOCAL m.n, m.lcCENTURY, m.liSele, m.loRam, m.lihFile
      STORE FCREATE(m.tcFile) TO m.lihFile, This._hFile 
      IF m.lihFile<=0
         RETURN .F.
      ENDIF
      
      This.CreateHeader()


      * Setting the default row heights
      =FWRITE(m.lihFile, This.WriteDefaultRowHeight())

      * Setting the default column widths
      =FWRITE(m.lihFile, This.WriteDefaultColumnWidth())

      * Setting the column forced widths
      IF TYPE("This.aColWidths(1, 1)") = "N"
         FOR m.n = 1 TO ALEN(This.aColWidths, 1)
             =FWRITE(m.lihFile, This.WriteColumnInfoRecord(This.aColWidths(m.n, 1), This.aColWidths(m.n, 2)))
         ENDFOR 
      ENDIF 

      m.lcCENTURY=SET("CENTURY")
      IF This.cDateType="B"
         SET CENTURY ON
      ENDIF

      * Writing the cells
      m.liSele=SELECT()
      SELECT (This.cACells)
      SCAN ALL
           SCATTER NAME m.loRam
           =This.WriteCell(@m.loRam)
      ENDSCAN
      SELECT (m.liSele)
      
      SET CENTURY &lcCENTURY.

      =FWRITE(m.lihFile, BINTOC(EOFRecord, "4RS"))
      =FCLOSE(m.lihFile)
      This._hFile=0
      RETURN .T. 
   ENDFUNC 


   FUNCTION CreateHeader
      *
      * FoxyXLS::CreateHeader()
      *
      =FWRITE(This._hFile, This.GetMainHeader())
      =FWRITE(This._hFile, This.WriteAuthorRecord())
      =FWRITE(This._hFile, This.WriteCodepageRecord())
      =FWRITE(This._hFile, This.WriteFontTable())
      =FWRITE(This._hFile, This.WriteHeaderRecord())
      =FWRITE(This._hFile, This.WriteFooterRecord())
      =FWRITE(This._hFile, This.WriteFormatTable())
      =FWRITE(This._hFile, This.WriteWindowProtectRecord())
      =FWRITE(This._hFile, This.WriteXFTable())
      =FWRITE(This._hFile, This.WriteStyleTable())
   ENDFUNC 


   FUNCTION GetMainHeader
      *
      * FoxyXLS::GetMainHeader()
      *
      LOCAL m.lcBiffHeader
*!*       m.lcBiffHeader = CHR(0x09)+CHR(0x02)+; && BIFF version
*!*                        CHR(0x08)+CHR(0x00)+; && Legth
*!*                        CHR(0x00)+CHR(0x00)+; && Not used (BIFF version)
*!*                        CHR(0x10)+CHR(0x00)+; && Sheet
*!*                        CHR(0x00)+CHR(0x00)+; && Not used (Build identifier - BIFF 5,8)
*!*                        CHR(0x00)+CHR(0x00)   && Not used (Build year - BIFF 5,8)
      m.lcBiffHeader = CHR(0x09)+CHR(0x02)+; && BIFF version
                       CHR(0x06)+CHR(0x00)+; && Legth
                       CHR(0x00)+CHR(0x00)+; && Not used (BIFF version)
                       CHR(0x10)+CHR(0x00)+; && Sheet
                       CHR(0x00)+CHR(0x00) && Not used
      RETURN m.lcBiffHeader
   ENDFUNC 


   FUNCTION WriteAuthorRecord
      *
      * FoxyXLS::WriteAuthorRecord()
      *
      LOCAL m.lcBiffAuthor, m.lcPom
      m.lcPom=LEFT(This.cAuthor,31)
      m.lcBiffAuthor = CHR(0x5C)+CHR(0x00)+CHR(0x20)+CHR(0x00)+CHR(LEN(m.lcPom))+PADR(m.lcPom, 31, " ")  && 31 positions, but can be bigger
      RETURN m.lcBiffAuthor
   ENDFUNC 


   FUNCTION WriteCodepageRecord
      *
      * FoxyXLS::WriteCodepageRecord()
      *
      LOCAL m.lcBiffCodePage
      IF TYPE("This.nCodePage") = "C"
         This.nCodePage = VAL(This.nCodePage)
      ENDIF  
      
      m.lcBiffCodePage = CHR(0x42)+CHR(0x00)+CHR(0x02)+CHR(0x00)+ BINTOC(This.nCodePage, "2RS")
      *1250 = E204 (little endian)
      RETURN m.lcBiffCodePage
   ENDFUNC 


   FUNCTION WriteFontTable
      *
      * FoxyXLS::WriteFontTable()
      *
      * Write font table
      LOCAL m.lcBiffFontRecord, m.n
      m.lcBiffFontRecord = ""
      FOR m.n = 1 TO ALEN(This.aFonts, 1)
          IF NOT EMPTY(This.aFonts(m.n, 1))
             m.lcBiffFontRecord = m.lcBiffFontRecord + This.WriteFontRecord(m.n)
          ENDIF 
      ENDFOR 
      RETURN m.lcBiffFontRecord
   ENDFUNC 


   FUNCTION WriteFontRecord
      *
      * FoxyXLS::WriteFontRecord()
      *
      LPARAMETERS m.tnIndex

      * Getting the correct Forecolor value
      LOCAL m.lcFont,m.lcStyle, m.lcArray, m.lnIntStyle
      m.lcFont  = This.aFonts(m.tnIndex, 2)
      m.lcStyle = This.aFonts(m.tnIndex, 4)

      m.lnIntStyle = IIF("B" $ m.lcStyle, 1, 0) + ;
                     IIF("I" $ m.lcStyle, 2, 0) + ;
                     IIF("U" $ m.lcStyle, 4, 0) + ;
                     IIF("S" $ m.lcStyle, 8, 0)

      m.lcArray = BINTOC(FontRecord, "2RS")   + ;                      && 0h3102 = FontRecord
                  BINTOC(LEN(m.lcFont) + 7, "2RS") + ;                 && Fontname length + 7
                  BINTOC(This.aFonts(m.tnIndex, 3) * 20, "2RS")    + ; && FontSize
                  BINTOC(m.lnIntStyle, "2RS")     + ;                  && FontStyle
                  BINTOC(This.aFonts(m.tnIndex, 6)   , "2RS")   + ;    && Color
                  CHR(LEN(m.lcFont))              + ;                  && Fontname length
                  m.lcFont                                             && Fontname

      RETURN m.lcArray
   ENDFUNC


   FUNCTION WriteHeaderRecord
      *
      * FoxyXLS::WriteHeaderRecord()
      *
      LOCAL m.lcBiffHeaderRecord
      m.lcBiffHeaderRecord = BINTOC(HeaderRecord, "4RS")
      RETURN m.lcBiffHeaderRecord
   ENDFUNC 


   FUNCTION WriteFooterRecord
      *
      * FoxyXLS::WriteFooterRecord()
      *
      LOCAL m.lcBiffFooterRecord
      m.lcBiffFooterRecord = BINTOC(FooterRecord, "4RS")
      RETURN m.lcBiffFooterRecord
   ENDFUNC 


   FUNCTION WriteFormatTable
      *
      * FoxyXLS::WriteFormatTable()
      *
      * Creating the table array
      LOCAL m.lcBiffFormatRecord
      m.lcBiffFormatRecord = ""

      FOR m.n = 1 TO ALEN(This.aFmts, 1)
          m.lcBiffFormatRecord = m.lcBiffFormatRecord + This.WriteFormatRecord(This.aFmts(m.n))
      ENDFOR
      RETURN m.lcBiffFormatRecord
   ENDFUNC 


   FUNCTION WriteFormatRecord
      *
      * FoxyXLS::WriteFormatRecord()
      *
      LPARAMETERS m.tcFormat
      * #DEFINE FormatRecord         0x001E
      LOCAL m.lcReturn
      m.lcReturn = BINTOC(FormatRecord, "2RS") + ;        && formatrecord
                   BINTOC(LEN(m.tcFormat) + 1,"2RS") + ;  && length block
                   CHR(LEN(m.tcFormat)) + ;               && format length
                   m.tcFormat                             && format
      RETURN m.lcReturn 
   ENDFUNC 


   FUNCTION WriteWindowProtectRecord
      *
      * FoxyXLS::WriteWindowProtectRecord()
      *
      * #DEFINE WindowProtectRecord  0x0019
      LOCAL m.lcBiffWindowProtectRecord
      m.lcBiffWindowProtectRecord = BINTOC(WindowProtectRecord, "2RS")
      RETURN m.lcBiffWindowProtectRecord
   ENDFUNC 


   FUNCTION WriteXFTable
      *
      * FoxyXLS::WriteXFTable()
      *
      * #DEFINE XFRecord             0x0243
      LOCAL m.lcBiffXFTable, m.lcAuxTable, m.lcArrayTable, m.lnItem, m.n
      LOCAL m.laArrayTable(1)

      m.lcAuxTable =CHRTRAN(;
                   "0x0243,0x000C,0x0000,0x03F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0001,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0001,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0002,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0002,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0xF7F5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0000,0x0001,0x0000,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x2101,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x1F01,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x2001,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x1E01,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000," + ;
                   "0x0243,0x000C,0x0901,0xFBF5,0xFFF0,0xCE00,0x0000,0x0000", ",", CHR(10))

      m.lcArrayTable = ""
      FOR m.n = 1 TO ALINES(m.laArrayTable, m.lcAuxTable)
          m.lnItem = EVALUATE(m.laArrayTable(m.n))
          m.lcArrayTable =  m.lcArrayTable + LEFT(BINTOC(m.lnItem, "4RS"), 2)
      ENDFOR

      m.lcBiffXFTable=BINTOC(2, "4RS") + m.lcArrayTable
      
      m.lcBiffFX = This.CreateFxTable()
      RETURN m.lcBiffXFTable + m.lcBiffFx
   ENDFUNC 


   FUNCTION CreateFxTable
      *
      * FoxyXLS::CreateFxTable()
      *
      LOCAL m.lcTable, m.liSele, m.loRam
      m.lcTable = ""
      m.liSele=SELECT()
      SELECT (This.cAFXStyles)
      SET ORDER TO
      SCAN ALL
           SCATTER NAME m.loRam
           m.lcTable = m.lcTable + This.WriteFXRecord(@m.loRam)
      ENDSCAN
      SET ORDER TO "I01"
      SELECT (m.liSele)
      RETURN m.lcTable
   ENDFUNC 


   FUNCTION WriteFXRecord
      *
      * FoxyXLS::WriteFXRecord()
      *
      LPARAMETERS m.loRam
      * #DEFINE XFRecord       0x0243

      LOCAL m.lcRecord, m.lnFontIndex, m.lnFormatIndex, m.lnHorizontalAlignment, m.lnForeColor, m.lnBackColor, m.lnAttr,;
            m.lqBackColor, m.lnXlsColor

      m.lnFontIndex           = m.loRam.FONTINDEX
      m.lnFormatIndex         = m.loRam.FORMATINDEX
      m.lnHorizontalAlignment = m.loRam.ALIGN && 2 = Center
      m.lnBackColor           = m.loRam.BACKCOLOR
      m.lnAttr                = 0
      
      m.lnAttr = BITOR(m.lnAttr, 0x08)
      IF m.lnHorizontalAlignment > 0 && Alignment.General
         m.lnAttr = BITOR(m.lnAttr, 0x10)
      ENDIF 
      IF m.lnBackColor > 0 && Backcolor
         m.lnAttr = BITOR(m.lnAttr, 0x40)
      ENDIF 

      IF (EMPTY(m.lnBackColor)) OR (TYPE("m.lnBackColor") <> "N")
         m.lqBackColor = CHR(0x00)+CHR(0xCE) && 0xCE00
      ELSE 
         m.lnXlsColor = This.aColors(m.lnBackColor,3)
         m.lqBackColor = LEFT(BINTOC(m.lnXLSColor, "4RS"),2)
      ENDIF 

      m.lcRecord = BINTOC(XFRecord, "2RS") + ; && Xfrecord
                   BINTOC(0x000C, "2RS") + ;   && block legth
                   CHR(m.lnFontIndex) + ;      && FontIndex
                   CHR(m.lnFormatIndex) + ;    && FormatIndex
                   CHR(1) + ;                  && XF Type and Cell Protection (01H - locked, 02H - hidden, 04H - 0 = Cell XF; 1 = Style XF)
                   CHR(m.lnAttr) + ;           && Attributes Used from Parent Style XF
                   BINTOC(m.lnHorizontalAlignment, "2RS") + ; && Horizontal Alignment
                   m.lqBackColor + ;           && Backcolor
                   BINTOC(0x0000, "2RS") + ;   && Cell border lines (top-left)
                   BINTOC(0x0000, "2RS")       && Cell border lines (bottom-right)

      RETURN m.lcRecord
   ENDFUNC 


   FUNCTION WriteStyleTable
      *
      * FoxyXLS::WriteStyleTable()
      *
      * #DEFINE StyleRecord          0x0293   Linha 4D0
      LOCAL m.lcBiffStyle

      m.lcBiffStyle = BINTOC(StyleRecord, "2RS") + CHR(0x04)+CHR(0x00) + CHR(0x10)+CHR(0x80)+CHR(0x03)+CHR(0xFF) + ;
                      BINTOC(StyleRecord, "2RS") + CHR(0x0C)+CHR(0x00) + CHR(0x11)+CHR(0x00)+CHR(0x09)+CHR(0x43)+CHR(0x6F)+CHR(0x6D)+CHR(0x6D)+CHR(0x61)+CHR(0x20)+CHR(0x5B)+CHR(0x30)+CHR(0x5D) + ;
                      BINTOC(StyleRecord, "2RS") + CHR(0x04)+CHR(0x00) + CHR(0x12)+CHR(0x80)+CHR(0x04)+CHR(0xFF) + ; 
                      BINTOC(StyleRecord, "2RS") + CHR(0x0F)+CHR(0x00) + CHR(0x13)+CHR(0x00)+CHR(0x0C)+CHR(0x43)+CHR(0x75)+CHR(0x72)+CHR(0x72)+CHR(0x65)+CHR(0x6E)+CHR(0x63)+CHR(0x79)+CHR(0x20)+CHR(0x5B)+CHR(0x30)+CHR(0x5D) + ;
                      BINTOC(StyleRecord, "2RS") + CHR(0x04)+CHR(0x00) + CHR(0x00)+CHR(0x80)+CHR(0x00)+CHR(0xFF) + ;
                      BINTOC(StyleRecord, "2RS") + CHR(0x04)+CHR(0x00) + CHR(0x14)+CHR(0x80)+CHR(0x05)+CHR(0xFF)

      RETURN m.lcBiffStyle
   ENDFUNC 


   FUNCTION WriteDefaultRowHeight
      *
      * FoxyXLS::WriteDefaultRowHeight()
      *
      * #DEFINE DefRowHeight          0x0225
      LOCAL m.lqReturn
      m.lqReturn = ""
      IF NOT EMPTY(This.nDefaultRowHeight)
         m.lqReturn = BINTOC(DefRowHeight, "2RS") + ; && DefaultRowHeight record
                      BINTOC(4, "2RS") + ;            && block length
                      BINTOC(1, "2RS") + ;            && Option flags:
                      BINTOC(This.nDefaultRowHeight * 20, "2RS") && row height
      ENDIF 
      RETURN m.lqReturn

   *!* 5.31 DEFAULTROWHEIGHT
   *!* BIFF2 BIFF3 BIFF4 BIFF5 BIFF8
   *!* 0025H 0225H 0225H 0225H 0225H
   *!* This record specifies the default height and default flags for rows that do not have a corresponding ROW record
   *!* (?5.88).
   *!* Record DEFAULTROWHEIGHT, BIFF2:
   *!* Offset Size Contents
   *!* 0 2 Default height for unused rows, in twips = 1/20 of a point
   *!* Bit Mask Contents
   *!* 0-14 7FFFH Default height for unused rows, in twips = 1/20 of a point
   *!* 15 8000H 1 = Row height not changed manually
   *!* Record DEFAULTROWHEIGHT, BIFF3-BIFF8:
   *!* Offset Size Contents
   *!* 0 2 Option flags:
   *!*           Bit Mask Contents
   *!*           0 0001H 1 = Row height and default font height do not match
   *!*           1 0002H 1 = Row is hidden
   *!*           2 0004H 1 = Additional space above the row
   *!*           3 0008H 1 = Additional space below the row
   *!* 2 2 Default height for unused rows, in twips = 1 /20 of a point
   ENDFUNC 


   FUNCTION WriteDefaultColumnWidth
      *
      * FoxyXLS::WriteDefaultColumnWidth()
      *
      * #DEFINE DefColWidth          0x0055
      LOCAL m.lqReturn
      m.lqReturn = ""
      IF NOT EMPTY(This.nDefaultColumnWidth)
         m.lqReturn = BINTOC(DefColWidth, "2RS") + ;          && DefaultColumnWidth
                      BINTOC(2, "2RS") + ;                    && block length
                      BINTOC(This.nDefaultColumnWidth, "2RS") && width
      ENDIF 
      RETURN m.lqReturn

   *!* 5.32 DEFCOLWIDTH
   *!* BIFF2 BIFF3 BIFF4 BIFF5 BIFF8
   *!* 0055H 0055H 0055H 0055H 0055H
   *!* This record specifies the default column width for columns that do not have a specific width set using the records
   *!* COLWIDTH (BIFF2, ?5.20), COLINFO (BIFF3-BIFF8, ?5.18), or STANDARDWIDTH (?5.101).
   *!* Record DEFCOLWIDTH, BIFF2-BIFF8:
   *!* Offset Size Contents
   *!* 0 2 Column width in characters, using the width of the zero character from default font (first
   *!* FONT record in the file). Excel adds some extra space to the default width, depending on
   *!* the default font and default font size. The algorithm how to exactly calculate the resulting
   *!* column width is not known.
   *!* Example: The default width of 8 set in this record results in a column width of
   *!* 8.43 using Arial font with a size of 10 points.

   ENDFUNC 


   FUNCTION WriteColumnInfoRecord
      *
      * FoxyXLS::WriteColumnInfoRecord()
      *
      LPARAMETERS m.tnCol, m.tnWidth
      * #DEFINE ColumnInfoRecord     0x007D

      m.tnCol = MAX(0, m.tnCol - 1)
         
      LOCAL m.lcReturn
      m.lcReturn = BINTOC(ColumnInfoRecord, "2RS") + ;       && ColumnInfo Record
                   BINTOC(12, "2RS") + ;                     && block length
                   BINTOC(m.tnCol, "2RS") + ;                && Index to first column in the range
                   BINTOC(m.tnCol, "2RS") + ;                && Index to last column in the range
                   BINTOC((m.tnWidth * 256 / 7), "2RS") + ;  && Column width
                   BINTOC(15, "2RS") + ;                     && Index to XF record 
                   BINTOC(0, "2RS") + ;                      && Option flags
                   BINTOC(0, "2RS")                          && Not used
      RETURN m.lcReturn

*!*       * Setting the default row heights
*!*       * lcBiffCol = lcBiffCol + ;
*!*        0h0802 + ;  && Index
*!*        0h1000 + ;  && Size
*!*        0h0000 + ;  && Row number
*!*        0h0000 + ;  && Index to column of the first cell which is described by a cell record
*!*        0h0400 + ;  && Index to column of the last cell which is described by a cell record, increased by 1
*!*        BINTOC(40 * 20, "2RS") + ;  && Height of the row, in twips = 1/20 of a point
*!*        0h0000 + ;  && Not used
*!*        0h1022 + ;  && a relative offset to calculate stream position of the first cell record for this row (?4.7.2)
*!*        0h0100f000  && Option flags and default row formatting
   ENDFUNC


   FUNCTION WriteCell
      *
      * FoxyXLS::WriteCell()
      *
      LPARAMETERS m.loRam
      LOCAL m.lcType, m.tuValue
      
      #IF VFP_VERSION<9
       m.tuValue=IIF(m.loRam.AGROUP=1, This.aCells(m.loRam.AINDEX), EVALUATE("This.aCells"+LTRIM(STR(m.loRam.AGROUP,2))+"(m.loRam.AINDEX)") )
      #ELSE
       m.tuValue=This.aCells(m.loRam.AINDEX)
      #ENDIF
      
      m.lcType = TYPE("m.tuValue")
      DO CASE
         CASE ISNULL(m.tuValue)
              =FWRITE(This._hFile, This.WriteEmpty(m.loRam))
              
         CASE m.lcType = "C"
              =FWRITE(This._hFile, This.WriteString(m.loRam, m.tuValue))

         CASE m.lcType = "L"
              =FWRITE(This._hFile, This.WriteBoolean(m.loRam, m.tuValue))

         CASE m.lcType = "N" AND INT(m.tuValue)=m.tuValue
              =FWRITE(This._hFile, This.WriteInt(m.loRam, m.tuValue))

         CASE m.lcType = "N"
              =FWRITE(This._hFile, This.WriteDoble(m.loRam, m.tuValue))

         CASE m.lcType = "D" AND m.tuValue>={^1900-01-01}
              =FWRITE(This._hFile, This.WriteDate(m.loRam, m.tuValue))

         CASE m.lcType = "D" AND m.tuValue<{^1900-01-01}
              =FWRITE(This._hFile, This.WriteString(m.loRam, DTOC(m.tuValue)))

         CASE m.lcType = "T" AND m.tuValue>={^1900-01-01}
              =FWRITE(This._hFile, This.WriteDateTime(m.loRam, m.tuValue))

         CASE m.lcType = "T" AND m.tuValue<{^1900-01-01}
              =FWRITE(This._hFile, This.WriteString(m.loRam, TTOC(m.tuValue)))

         OTHERWISE
              =FWRITE(This._hFile, This.WriteEmpty(m.loRam))

      ENDCASE
      RETURN 
   ENDFUNC


   FUNCTION WriteString
      *
      * FoxyXLS::WriteString()
      *
      LPARAMETERS m.loRam, m.tcString
      IF LEN(m.tcString) > 255
         m.tcString = LEFT(m.tcString, 255)
      ENDIF 
      * #DEFINE LabelRecord          0x0204
      RETURN BINTOC(LabelRecord, "2RS") + BINTOC(8+LEN(m.tcString), "2RS") + ;       && Header + block length
             BINTOC(m.loRam.ROWNO - 1, "2RS") + BINTOC(m.loRam.COLNO - 1, "2RS") + ; && Row, Col
             BINTOC(m.loRam.FXSTYLE, "2RS") + ;                                      && (ushort)cell.FXIndex
             BINTOC(LEN(m.tcString), "2RS") + m.tcString
   ENDFUNC


   FUNCTION WriteBoolean
      *
      * FoxyXLS::WriteBoolean()
      *
      LPARAMETERS m.loRam, m.tnValue
      * #DEFINE BooleanRecord          0x0205
      RETURN BINTOC(BooleanRecord, "2RS") + BINTOC(8, "2RS") + ;                     && Header + block length
             BINTOC(m.loRam.ROWNO - 1, "2RS") + BINTOC(m.loRam.COLNO - 1, "2RS") + ; && Row, Col
             BINTOC(m.loRam.FXSTYLE, "2RS") + ;                                      && (ushort)cell.FXIndex
             IIF(m.tnValue, CHR(1), CHR(0))+CHR(0)
   ENDFUNC
    

   FUNCTION WriteInt
      *
      * FoxyXLS::WriteInt()
      *
      LPARAMETERS m.loRam, m.tnValue
      
      LOCAL m.lnValue
      m.lnValue = BITLSHIFT(m.tnValue, 2) + 2

      * #DEFINE IntegerRecord          0x027E
      RETURN BINTOC(IntegerRecord, "2RS") + BINTOC(10, "2RS") + ;                   && Header + block length
             BINTOC(m.loRam.ROWNO - 1, "2RS") + BINTOC(m.loRam.COLNO - 1, "2RS") +; && Row, Col
             BINTOC(m.loRam.FXSTYLE, "2RS") + BINTOC(m.lnValue, "4RS")              && (ushort)cell.FXIndex
   ENDFUNC


   FUNCTION WriteDoble 
      *
      * FoxyXLS::WriteDoble()
      *
      LPARAMETERS m.loRam, m.tnValue
      
      * #DEFINE NumberRecord         0x0203
      RETURN BINTOC(NumberRecord, "2RS") + BINTOC(14, "2RS")     + ;                 && Header + block length
             BINTOC(m.loRam.ROWNO - 1, "2RS") + BINTOC(m.loRam.COLNO - 1, "2RS") + ; && Row, Col
             BINTOC(m.loRam.FXSTYLE, "2RS")                           + ;            && (ushort)cell.FXIndex
             BINTOC(m.tnValue, "8SB")
   ENDFUNC


   FUNCTION WriteDate
      *
      * FoxyXLS::WriteDate()
      *
      LPARAMETERS m.loRam, m.tdDate
      
      * #DEFINE NumberRecord         0x0203
      LOCAL m.lnDateValue, m.lcDateFormat
      m.lnDateValue = m.tdDate-{^1899-12-30}-IIF(m.tdDate<={^1900-02-28},1,0)
      RETURN BINTOC(NumberRecord, "2RS") + BINTOC(14, "2RS") + ;                     && Header + block length
             BINTOC(m.loRam.ROWNO - 1, "2RS") + BINTOC(m.loRam.COLNO - 1, "2RS") + ; && Row, Col
             BINTOC(m.loRam.FXSTYLE, "2RS") + ;                                      && (ushort)cell.FXIndex
             BINTOC(m.lnDateValue, "8SB")
   ENDFUNC


   FUNCTION WriteDatetime
      *
      * FoxyXLS::WriteDatetime()
      *
      LPARAMETERS m.loRam, m.ttDateTime
      
      * #DEFINE NumberRecord         0x0203
      LOCAL m.lnDateValue, m.lcDateFormat
      m.lnDateValue=(m.ttDateTime-{^1899-12-30 00:00:00}-IIF(m.ttDateTime<={^1900-02-28},86400,0))/86400.00000
      RETURN BINTOC(NumberRecord, "2RS") + BINTOC(14, "2RS") + ;                     && Header + block length
             BINTOC(m.loRam.ROWNO - 1, "2RS") + BINTOC(m.loRam.COLNO - 1, "2RS") + ; && Row, Col
             BINTOC(m.loRam.FXSTYLE, "2RS") + ;                                      && (ushort)cell.FXIndex
             BINTOC(m.lnDateValue, "8SB")
   ENDFUNC


   FUNCTION WriteEmpty
      *
      * FoxyXLS::WriteEmpty()
      *
      LPARAMETERS m.loRam
      RETURN BINTOC(EmptyRecord, "2RS") + BINTOC(6, "2RS")+ ;                        && Header + block length
             BINTOC(m.loRam.ROWNO - 1, "2RS") + BINTOC(m.loRam.COLNO - 1, "2RS") + ; && Row, Col
             BINTOC(m.loRam.FXSTYLE, "2RS")                                          && (ushort)cell.FXIndex
   ENDFUNC

ENDDEFINE