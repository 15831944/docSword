using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Collections;

using System.Xml;
using System.Xml.Serialization;


namespace OfficeTools.Common
{
    // 参照Word.Font的对象
    [Serializable]
    [XmlType]
    public class ClassFont /*: Font*/
    {
        // [XmlIgnoreAttribute]
        public enum euMembers
        {
            Zero = 0,
            AllCaps,
            Animation,
            Bold,
            BoldBi,
            Borders,
            Color,
            ColorIndex,
            ColorIndexBi,
            DiacriticColor,
            DisableCharacterSpaceGrid,
            DoubleStrikeThrough,
            Duplicate,
            Emboss,
            EmphasisMark,
            Engrave,
            Hidden,
            Italic,
            ItalicBi,
            Kerning,
            Name,
            NameAscii,
            NameBi,
            NameFarEast,
            Outline,
            Position,
            Scaling,
            Shading,
            Shadow,
            Size,
            SizeBi,
            SmallCaps,
            Spacing,
            StrikeThrough,
            Subscript,
            Superscript,
            Underline,
            UnderlineColor,
        };

        [XmlIgnoreAttribute]
        public HashSet<int> setMembers = null;

        [XmlAttribute]
        public int AllCaps;// { get; set; }
        [XmlAttribute]
        public WdAnimation Animation;// { get; set; }
        // public Application Application;// { get; }
        [XmlAttribute]
        public int Bold;// { get; set; }
        [XmlAttribute]
        public int BoldBi;// { get; set; }
        // 对应Word.Borders的对象
        [XmlIgnoreAttribute]
        public ClassBorders Borders = new ClassBorders();
        //public Borders Borders;// { get; set; }
        [XmlAttribute]
        public WdColor Color;// { get; set; }
        [XmlAttribute]
        public WdColorIndex ColorIndex;// { get; set; }
        [XmlAttribute]
        public WdColorIndex ColorIndexBi;// { get; set; }
        // public int Creator;// { get; }
        [XmlAttribute]
        public WdColor DiacriticColor;// { get; set; }
        [XmlAttribute]
        public bool DisableCharacterSpaceGrid;// { get; set; }
        [XmlAttribute]
        public int DoubleStrikeThrough;// { get; set; }

        [XmlIgnoreAttribute]
        public Font Duplicate;// { get; }
        [XmlAttribute]
        public int Emboss;// { get; set; }
        [XmlAttribute]
        public WdEmphasisMark EmphasisMark;// { get; set; }
        [XmlAttribute]
        public int Engrave;// { get; set; }
        [XmlAttribute]
        public int Hidden;// { get; set; }
        [XmlAttribute]
        public int Italic;// { get; set; }
        [XmlAttribute]
        public int ItalicBi;// { get; set; }
        [XmlAttribute]
        public float Kerning;// { get; set; }
        [XmlAttribute]
        public string Name = "";// { get; set; }
        [XmlAttribute]
        public string NameAscii = "";// { get; set; }
        [XmlAttribute]
        public string NameBi = "";// { get; set; }
        [XmlAttribute]
        public string NameFarEast = "";// { get; set; }
        //public string NameOther = "";// { get; set; }
        [XmlAttribute]
        public int Outline;// { get; set; }
        // public dynamic Parent { get; }
        [XmlAttribute]
        public int Position;// { get; set; }
        [XmlAttribute]
        public int Scaling;// { get; set; }

        [XmlIgnoreAttribute]
        public ClassShading Shading = new ClassShading();
        // public Shading Shading;// { get; }
        [XmlAttribute]
        public int Shadow;// { get; set; }
        [XmlAttribute]
        public float Size;// { get; set; }
        [XmlAttribute]
        public float SizeBi;// { get; set; }
        [XmlAttribute]
        public int SmallCaps;// { get; set; }
        [XmlAttribute]
        public float Spacing;// { get; set; }
        [XmlAttribute]
        public int StrikeThrough;// { get; set; }
        [XmlAttribute]
        public int Subscript;// { get; set; }
        [XmlAttribute]
        public int Superscript;// { get; set; }
        [XmlAttribute]
        public WdUnderline Underline;// { get; set; }
        [XmlAttribute]
        public WdColor UnderlineColor;// { get; set; }


        public Boolean isSet(ClassFont.euMembers enMem)
        {
            if(setMembers != null)
            {
                return setMembers.Contains((int)enMem);
            }

            return false;
        }

        public void ClearSelMember()
        {
            if (setMembers == null)
            {
                setMembers = new HashSet<int>();
            }
            else
            {
                setMembers.Clear();
            }

            return;
        }


        public Boolean AddSelMember(int nFontMember)
        {
            if (setMembers == null)
            {
                setMembers = new HashSet<int>();
            }

            Boolean bRet = setMembers.Add(nFontMember);

            return bRet;
        }


        public Boolean RemoveSelMember(int nFontMember)
        {
            if (setMembers == null)
            {
                setMembers = new HashSet<int>();
            }

            Boolean bRet = setMembers.Remove(nFontMember);

            return bRet;
        }

        public int SelCopy2(Word.Font wFnt)
        {
            if (setMembers == null)
            {
                return 0;
            }

            euMembers euItem = euMembers.Zero;

            foreach(int nItem in setMembers)
            {
                euItem = (euMembers)nItem;

                switch (euItem)
                {
                    case euMembers.AllCaps:
                        wFnt.AllCaps = AllCaps;
                        break;

                    case euMembers.Animation:
                        wFnt.Animation = Animation;
                        break;

                    case euMembers.Bold:
                        wFnt.Bold = Bold;
                        break;

                    case euMembers.BoldBi:
                        wFnt.BoldBi = BoldBi;
                        break;

                    case euMembers.Borders:
                        // wFnt.Borders = Borders;
                        break;

                    case euMembers.Color:
                        wFnt.Color = Color;
                        break;

                    case euMembers.ColorIndex:
                        wFnt.ColorIndex = ColorIndex;
                        break;

                    case euMembers.ColorIndexBi:
                        wFnt.ColorIndexBi = ColorIndexBi;
                        break;

                    case euMembers.DiacriticColor:
                        wFnt.DiacriticColor = DiacriticColor;
                        break;

                    case euMembers.DisableCharacterSpaceGrid:
                        wFnt.DisableCharacterSpaceGrid = DisableCharacterSpaceGrid;
                        break;

                    case euMembers.DoubleStrikeThrough:
                        wFnt.DoubleStrikeThrough = DoubleStrikeThrough;
                        break;

                    case euMembers.Duplicate:
                        //wFnt.Duplicate = Duplicate;
                        break;

                    case euMembers.Emboss:
                        wFnt.Emboss = Emboss;
                        break;

                    case euMembers.EmphasisMark:
                        wFnt.EmphasisMark = EmphasisMark;
                        break;

                    case euMembers.Engrave:
                        wFnt.Engrave = Engrave;
                        break;

                    case euMembers.Hidden:
                        wFnt.Hidden = Hidden;
                        break;

                    case euMembers.Italic:
                        wFnt.Italic = Italic;
                        break;

                    case euMembers.ItalicBi:
                        wFnt.ItalicBi = ItalicBi;
                        break;

                    case euMembers.Kerning:
                        wFnt.Kerning = Kerning;
                        break;

                    case euMembers.Name:
                        wFnt.Name = Name;
                        break;

                    case euMembers.NameAscii:
                        wFnt.NameAscii = NameAscii;
                        break;

                    case euMembers.NameBi:
                        wFnt.NameBi = NameBi;
                        break;

                    case euMembers.NameFarEast:
                        wFnt.NameFarEast = NameFarEast;
                        break;

                    case euMembers.Outline:
                        wFnt.Outline = Outline;
                        break;

                    case euMembers.Position:
                        wFnt.Position = Position;
                        break;

                    case euMembers.Scaling:
                        wFnt.Scaling = Scaling;
                        break;

                    case euMembers.Shading:
                        //wFnt.Shading = Shading;
                        break;

                    case euMembers.Shadow:
                        wFnt.Shadow = Shadow;
                        break;

                    case euMembers.Size:
                        wFnt.Size = Size;
                        break;

                    case euMembers.SizeBi:
                        wFnt.SizeBi = SizeBi;
                        break;

                    case euMembers.SmallCaps:
                        wFnt.SmallCaps = SmallCaps;
                        break;

                    case euMembers.Spacing:
                        wFnt.Spacing = Spacing;
                        break;

                    case euMembers.StrikeThrough:
                        wFnt.StrikeThrough = StrikeThrough;
                        break;

                    case euMembers.Subscript:
                        wFnt.Subscript = Subscript;
                        break;

                    case euMembers.Superscript:
                        wFnt.Superscript = Superscript;
                        break;

                    case euMembers.Underline:
                        wFnt.Underline = Underline;
                        break;

                    case euMembers.UnderlineColor:
                        wFnt.UnderlineColor = UnderlineColor;
                        break;

                    default:
                        break;
                }
            }

            return setMembers.Count; // sel
        }


        public int SelCopy2(ClassFont wFnt)
        {
            if (setMembers == null)
            {
                return 0;
            }

            wFnt.ClearSelMember();

            euMembers euItem = euMembers.Zero;

            foreach (int nItem in setMembers)
            {
                wFnt.AddSelMember(nItem);

                euItem = (euMembers)nItem;
                switch (euItem)
                {
                    case euMembers.AllCaps:
                        wFnt.AllCaps = AllCaps;
                        break;

                    case euMembers.Animation:
                        wFnt.Animation = Animation;
                        break;

                    case euMembers.Bold:
                        wFnt.Bold = Bold;
                        break;

                    case euMembers.BoldBi:
                        wFnt.BoldBi = BoldBi;
                        break;

                    case euMembers.Borders:
                        // wFnt.Borders = Borders;
                        break;

                    case euMembers.Color:
                        wFnt.Color = Color;
                        break;

                    case euMembers.ColorIndex:
                        wFnt.ColorIndex = ColorIndex;
                        break;

                    case euMembers.ColorIndexBi:
                        wFnt.ColorIndexBi = ColorIndexBi;
                        break;

                    case euMembers.DiacriticColor:
                        wFnt.DiacriticColor = DiacriticColor;
                        break;

                    case euMembers.DisableCharacterSpaceGrid:
                        wFnt.DisableCharacterSpaceGrid = DisableCharacterSpaceGrid;
                        break;

                    case euMembers.DoubleStrikeThrough:
                        wFnt.DoubleStrikeThrough = DoubleStrikeThrough;
                        break;

                    case euMembers.Duplicate:
                        //wFnt.Duplicate = Duplicate;
                        break;

                    case euMembers.Emboss:
                        wFnt.Emboss = Emboss;
                        break;

                    case euMembers.EmphasisMark:
                        wFnt.EmphasisMark = EmphasisMark;
                        break;

                    case euMembers.Engrave:
                        wFnt.Engrave = Engrave;
                        break;

                    case euMembers.Hidden:
                        wFnt.Hidden = Hidden;
                        break;

                    case euMembers.Italic:
                        wFnt.Italic = Italic;
                        break;

                    case euMembers.ItalicBi:
                        wFnt.ItalicBi = ItalicBi;
                        break;

                    case euMembers.Kerning:
                        wFnt.Kerning = Kerning;
                        break;

                    case euMembers.Name:
                        wFnt.Name = Name;
                        break;

                    case euMembers.NameAscii:
                        wFnt.NameAscii = NameAscii;
                        break;

                    case euMembers.NameBi:
                        wFnt.NameBi = NameBi;
                        break;

                    case euMembers.NameFarEast:
                        wFnt.NameFarEast = NameFarEast;
                        break;

                    case euMembers.Outline:
                        wFnt.Outline = Outline;
                        break;

                    case euMembers.Position:
                        wFnt.Position = Position;
                        break;

                    case euMembers.Scaling:
                        wFnt.Scaling = Scaling;
                        break;

                    case euMembers.Shading:
                        //wFnt.Shading = Shading;
                        break;

                    case euMembers.Shadow:
                        wFnt.Shadow = Shadow;
                        break;

                    case euMembers.Size:
                        wFnt.Size = Size;
                        break;

                    case euMembers.SizeBi:
                        wFnt.SizeBi = SizeBi;
                        break;

                    case euMembers.SmallCaps:
                        wFnt.SmallCaps = SmallCaps;
                        break;

                    case euMembers.Spacing:
                        wFnt.Spacing = Spacing;
                        break;

                    case euMembers.StrikeThrough:
                        wFnt.StrikeThrough = StrikeThrough;
                        break;

                    case euMembers.Subscript:
                        wFnt.Subscript = Subscript;
                        break;

                    case euMembers.Superscript:
                        wFnt.Superscript = Superscript;
                        break;

                    case euMembers.Underline:
                        wFnt.Underline = Underline;
                        break;

                    case euMembers.UnderlineColor:
                        wFnt.UnderlineColor = UnderlineColor;
                        break;

                    default:
                        break;
                }
            }

            return setMembers.Count; // sel
        }
        // for Font Format Dialog
        //public int CharacterWidthGrid;
        //public int ColorDialog;
        //public String KerningMin = "";
        //public String PointsBi = "";

        public String encode2String()
        {
            String strRet = "";
            strRet += "[Font_Start:Font_Start]";

            strRet += "[Font_AllCaps:" + AllCaps + "]";
            strRet += "[Font_Animation:" + (int)Animation + "]";
            // strRet += "[Application:" + Application + "]";
            strRet += "[Font_Bold:" + Bold + "]";
            strRet += "[Font_BoldBi:" + BoldBi + "]";
            // strRet += "[Borders:" + Borders + "]";
            strRet += "[Font_ColorIndex:" + (int)ColorIndex + "]";
            strRet += "[Font_ColorIndexBi:" + (int)ColorIndexBi + "]";
            //strRet += "[ContextualAlternates:" + ContextualAlternates + "]";
            // strRet += "[Creator:" + Creator + "]";
            strRet += "[Font_DiacriticColor:" + (int)DiacriticColor + "]";
            strRet += "[Font_DisableCharacterSpaceGrid:" + DisableCharacterSpaceGrid + "]";
            strRet += "[Font_DoubleStrikeThrough:" + DoubleStrikeThrough + "]";
            // strRet += "[Duplicate:" + Duplicate + "]";
            strRet += "[Font_Emboss:" + Emboss + "]";
            strRet += "[Font_EmphasisMark:" + (int)EmphasisMark + "]";
            strRet += "[Font_Engrave:" + Engrave + "]";
            // strRet += "[Fill:" + Fill + "]";
            // strRet += "[Glow:" + Glow + "]";
            strRet += "[Font_Hidden:" + Hidden + "]";
            strRet += "[Font_Italic:" + Italic + "]";
            strRet += "[Font_ItalicBi:" + ItalicBi + "]";
            strRet += "[Font_Kerning:" + Kerning + "]";
            // ?strRet += "[Ligatures:" + (int)Ligatures + "]";
            // strRet += "[Line:" + Line + "]";
            strRet += "[Font_Name:" + Name + "]";
            strRet += "[Font_NameAscii:" + NameAscii + "]";
            strRet += "[Font_NameBi:" + NameBi + "]";
            strRet += "[Font_NameFarEast:" + NameFarEast + "]";
            // strRet += "[NameOther:" + NameOther + "]";
            //strRet += "[NumberForm:" + (int)NumberForm + "]";
            //strRet += "[NumberSpacing:" + (int)NumberSpacing + "]";
            strRet += "[Font_Outline:" + Outline + "]";
            // strRet += "[Parent:" + Parent + "]";
            strRet += "[Font_Position:" + Position + "]";
            // strRet += "[Reflection:" + Reflection + "]";
            strRet += "[Font_Scaling:" + Scaling + "]";
            // strRet += "[Shading:" + Shading + "]";
            strRet += "[Font_Shadow:" + Shadow + "]";
            strRet += "[Font_Size:" + Size + "]";
            strRet += "[Font_SizeBi:" + SizeBi + "]";
            strRet += "[Font_SmallCaps:" + SmallCaps + "]";
            strRet += "[Font_Spacing:" + Spacing + "]";
            strRet += "[Font_StrikeThrough:" + StrikeThrough + "]";
            //strRet += "[StylisticSet:" + (int)StylisticSet + "]";
            strRet += "[Font_Subscript:" + Subscript + "]";
            strRet += "[Font_Superscript:" + Superscript + "]";
            // strRet += "[TextColor:" + TextColor + "]";
            // strRet += "[TextShadow:" + TextShadow + "]";
            // strRet += "[ThreeD:" + ThreeD + "]";
            strRet += "[Font_Underline:" + (int)Underline + "]";
            strRet += "[Font_UnderlineColor:" + (int)UnderlineColor + "]";

            strRet += "[Font_End:Font_End]";

            return strRet;
        }

        public int decodeFromString(Hashtable hashFeatures)
        {
            if (hashFeatures == null || hashFeatures.Count == 0)
            {
                return 1;
            }

            int nDefaultVal = (int)Word.WdConstants.wdUndefined;// 进行赋值

            String strVal = "";
            int nVal = 0;
            float fVal = 0.0f;

            strVal = (String)hashFeatures["Font_AllCaps"];
            if (int.TryParse(strVal, out AllCaps))
            {
            }
            else
            {
                AllCaps = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Animation"];
            if (int.TryParse(strVal, out nVal))
            {
                Animation = (WdAnimation)nVal;
            }
            else
            {
                Animation = WdAnimation.wdAnimationNone;
            }

            strVal = (String)hashFeatures["Font_Bold"];
            if (int.TryParse(strVal, out Bold))
            {
            }
            else
            {
                Bold = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_BoldBi"];
            if (int.TryParse(strVal, out BoldBi))
            {
            }
            else
            {
                BoldBi = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Color"];
            if (int.TryParse(strVal, out nVal))
            {
                Color = (WdColor)nVal;
            }
            else
            {
                Color = WdColor.wdColorAutomatic;
            }

            strVal = (String)hashFeatures["Font_ColorIndex"];
            if (int.TryParse(strVal, out nVal))
            {
                ColorIndex = (WdColorIndex)nVal;
            }
            else
            {
                ColorIndex = WdColorIndex.wdAuto;
            }

            strVal = (String)hashFeatures["Font_ColorIndexBi"];
            if (int.TryParse(strVal, out nVal))
            {
                ColorIndexBi = (WdColorIndex)nVal;
            }
            else
            {
                ColorIndexBi = WdColorIndex.wdAuto;
            }

            // Creator = nInit;
            strVal = (String)hashFeatures["Font_DiacriticColor"];
            if (int.TryParse(strVal, out nVal))
            {
                DiacriticColor = (WdColor)nVal;
            }
            else
            {
                DiacriticColor = WdColor.wdColorAutomatic;
            }

            strVal = (String)hashFeatures["Font_DisableCharacterSpaceGrid"];
            if (Boolean.TryParse(strVal, out DisableCharacterSpaceGrid))
            {
                
            }
            else
            {
                DisableCharacterSpaceGrid = false;
            }

            strVal = (String)hashFeatures["Font_DoubleStrikeThrough"];
            if (int.TryParse(strVal, out nVal))
            {
                DoubleStrikeThrough = nVal;
            }
            else
            {
                DoubleStrikeThrough = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Emboss"];
            if (int.TryParse(strVal, out nVal))
            {
                Emboss = nVal;
            }
            else
            {
                Emboss = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_EmphasisMark"];
            if (int.TryParse(strVal, out nVal))
            {
                EmphasisMark = (WdEmphasisMark)nVal;
            }
            else
            {
                EmphasisMark = WdEmphasisMark.wdEmphasisMarkNone;
            }


            strVal = (String)hashFeatures["Font_Engrave"];
            if (int.TryParse(strVal, out nVal))
            {
                Engrave = nVal;
            }
            else
            {
                Engrave = nDefaultVal;
            }
            
            strVal = (String)hashFeatures["Font_Hidden"];
            if (int.TryParse(strVal, out nVal))
            {
                Hidden = nVal;
            }
            else
            {
                Hidden = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Italic"];
            if (int.TryParse(strVal, out nVal))
            {
                Italic = nVal;
            }
            else
            {
                Italic = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_ItalicBi"];
            if (int.TryParse(strVal, out nVal))
            {
                ItalicBi = nVal;
            }
            else
            {
                ItalicBi = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Kerning"];
            if (float.TryParse(strVal, out fVal))
            {
                Kerning = fVal;
            }
            else
            {
                Kerning = nDefaultVal;
            }

            Name = (String)hashFeatures["Font_Name"];// 进行赋值
            NameAscii = (String)hashFeatures["Font_NameAscii"];// 进行赋值
            NameBi = (String)hashFeatures["Font_NameBi"];// 进行赋值
            NameFarEast = (String)hashFeatures["Font_NameFarEast"];// 进行赋值
            //NameOther = "";// 进行赋值

            strVal = (String)hashFeatures["Font_Outline"];
            if (int.TryParse(strVal, out nVal))
            {
                Outline = nVal;
            }
            else
            {
                Outline = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Position"];
            if (int.TryParse(strVal, out nVal))
            {
                Position = nVal;
            }
            else
            {
                Position = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Scaling"];
            if (int.TryParse(strVal, out nVal))
            {
                Scaling = nVal;
            }
            else
            {
                Scaling = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Shadow"];
            if (int.TryParse(strVal, out nVal))
            {
                Shadow = nVal;
            }
            else
            {
                Shadow = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Size"];
            if (float.TryParse(strVal, out fVal))
            {
                Size = fVal;
            }
            else
            {
                Size = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_SizeBi"];
            if (float.TryParse(strVal, out fVal))
            {
                SizeBi = fVal;
            }
            else
            {
                SizeBi = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_SmallCaps"];
            if (int.TryParse(strVal, out nVal))
            {
                SmallCaps = nVal;
            }
            else
            {
                SmallCaps = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Spacing"];
            if (int.TryParse(strVal, out nVal))
            {
                Spacing = nVal;
            }
            else
            {
                Spacing = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_StrikeThrough"];
            if (int.TryParse(strVal, out nVal))
            {
                StrikeThrough = nVal;
            }
            else
            {
                StrikeThrough = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Subscript"];
            if (int.TryParse(strVal, out nVal))
            {
                Subscript = nVal;
            }
            else
            {
                Subscript = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Superscript"];
            if (int.TryParse(strVal, out nVal))
            {
                Superscript = nVal;
            }
            else
            {
                Superscript = nDefaultVal;
            }

            strVal = (String)hashFeatures["Font_Underline"];
            if (int.TryParse(strVal, out nVal))
            {
                Underline = (WdUnderline)nVal;
            }
            else
            {
                Underline = WdUnderline.wdUnderlineNone;
            }

            strVal = (String)hashFeatures["Font_UnderlineColor"];
            if (int.TryParse(strVal, out nVal))
            {
                UnderlineColor = (WdColor)nVal;
            }
            else
            {
                UnderlineColor = WdColor.wdColorAutomatic;
            }

            return 0;
        }

        public int decodeFromString(String strRet)
        {
            // 
            Hashtable hashFeatures = ClassOfficeCommon.Decode(strRet);

            if(hashFeatures == null || hashFeatures.Count == 0)
            {
                return 1;
            }

            int nRet = decodeFromString(hashFeatures);

            return nRet;
        }


        public ClassFont()
        {
            setDefault();
            return;
        }

        // 设置缺省初始值
        public void setDefault()
        {
            int nInit = (int)Word.WdConstants.wdUndefined;// 进行赋值

            AllCaps = nInit;// 进行赋值
            Animation = (WdAnimation)nInit;//.wdAnimationNone;// 进行赋值
            Bold = nInit;// 进行赋值
            BoldBi = nInit;// 进行赋值

            Color = (WdColor)nInit;//.wdColorAutomatic;// 进行赋值
            ColorIndex = (WdColorIndex)nInit;//.wdAuto;// 进行赋值
            ColorIndexBi = (WdColorIndex)nInit;//.wdAuto;// 进行赋值
            // Creator = nInit;
            DiacriticColor = (WdColor)nInit;//.wdColorAutomatic;// 进行赋值
            DisableCharacterSpaceGrid = false;// 进行赋值
            DoubleStrikeThrough = nInit;// 进行赋值
            Emboss = nInit;
            EmphasisMark = (WdEmphasisMark)nInit;//.wdEmphasisMarkNone;// 进行赋值

            Engrave = nInit;// 进行赋值
            Hidden = nInit;// 进行赋值
            Italic = nInit;// 进行赋值
            ItalicBi = nInit;// 进行赋值
            Kerning = (float)nInit;// 进行赋值

            Name = "";// 进行赋值
            NameAscii = "";// 进行赋值
            NameBi = "";// 进行赋值
            NameFarEast = "";// 进行赋值
            //NameOther = "";// 进行赋值
            Outline = nInit;// 进行赋值

            Position = nInit;// 进行赋值
            Scaling = nInit;// 进行赋值

            Shadow = nInit;// 进行赋值
            Size = (float)nInit;// 进行赋值
            SizeBi = (float)nInit;// 进行赋值

            SmallCaps = nInit;// 进行赋值
            Spacing = (float)nInit;// 进行赋值
            StrikeThrough = nInit;// 进行赋值
            Subscript = nInit;// 进行赋值
            Superscript = nInit;// 进行赋值
            Underline = (WdUnderline)nInit;//WdUnderline.wdUnderlineNone;// 进行赋值
            UnderlineColor = (WdColor)nInit; //.wdColorAutomatic;// 进行赋值

            // CharacterWidthGrid = 0;
            DisableCharacterSpaceGrid = false;// 进行赋值
            
            return;
        }

        // 比较2个类的值，是否异同
        public Boolean diff(Word.Font srcFont)
        {
            if (this.AllCaps != srcFont.AllCaps) // 比较AllCaps
                return false;

            if (this.Animation != srcFont.Animation)// 比较Animation
                return false;

            if (this.Bold != srcFont.Bold)// 比较Bold
                return false;

            if (this.BoldBi != srcFont.BoldBi)// 比较BoldBi
                return false;

            // copyBorders(srcFont.Borders, this.Borders);
            // this.Borders.clone(srcFont.Borders);

            if (this.Color != srcFont.Color)// 比较Color
                return false;

            if (this.ColorIndex != srcFont.ColorIndex)// 比较ColorIndex
                return false;

            if (this.ColorIndexBi != srcFont.ColorIndexBi)// 比较ColorIndexBi
                return false;

            if (this.DiacriticColor != srcFont.DiacriticColor)// 比较DiacriticColor
                return false;

            if (this.DisableCharacterSpaceGrid != srcFont.DisableCharacterSpaceGrid)// 比较DisableCharacterSpaceGrid
                return false;

            if (this.DoubleStrikeThrough != srcFont.DoubleStrikeThrough)// 比较DoubleStrikeThrough
                return false;

            if (this.Emboss != srcFont.Emboss)// 比较Emboss
                return false;

            if (this.EmphasisMark != srcFont.EmphasisMark)// 比较EmphasisMark
                return false;

            if (this.Engrave != srcFont.Engrave)// 比较Engrave
                return false;

            if (this.Hidden != srcFont.Hidden)// 比较Hidden
                return false;

            if (this.Italic != srcFont.Italic)// 比较Italic
                return false;

            if (this.ItalicBi != srcFont.ItalicBi)// 比较ItalicBi
                return false;

            if (this.Kerning != srcFont.Kerning)// 比较Kerning
                return false;

            if (this.Name != srcFont.Name)// 比较Name
                return false;

            if (this.NameAscii != srcFont.NameAscii)// 比较NameAscii
                return false;

            if (this.NameBi != srcFont.NameBi)// 比较NameBi
                return false;

            if (this.NameFarEast != srcFont.NameFarEast)// 比较NameFarEast
                return false;

            //            if (this.NameOther != srcFont.NameOther)// 比较NameOther
//                return false;

            if (this.Outline != srcFont.Outline)// 比较Outline
                return false;

            if (this.Position != srcFont.Position)// 比较Position
                return false;

            if (this.Scaling != srcFont.Scaling)// 比较Scaling
                return false;

            // this.Shading = srcFont.Shading;// 比较Shading
            // this.Shading.clone(srcFont.Shading);

            if (this.Shadow != srcFont.Shadow)// 比较Shadow
                return false;

            if (this.Size != srcFont.Size)// 比较Size
                return false;

            if (this.SizeBi != srcFont.SizeBi)// 比较SizeBi
                return false;

            if (this.SmallCaps != srcFont.SmallCaps)// 比较SmallCaps
                return false;

            if (this.Spacing != srcFont.Spacing)// 比较Spacing
                return false;

            if (this.StrikeThrough != srcFont.StrikeThrough)// 比较StrikeThrough
                return false;

            if (this.Subscript != srcFont.Subscript)// 比较Subscript
                return false;

            if (this.Superscript != srcFont.Superscript)// 比较Superscript
                return false;

            if (this.Underline != srcFont.Underline)// 比较Underline
                return false;

            if (this.UnderlineColor != srcFont.UnderlineColor)// 比较UnderlineColor
                return false;

            return true;
        }

        // 比较2个类的值，是否异同
        public Boolean diff(ClassFont srcFont)
        {
            if (this.AllCaps != srcFont.AllCaps) // 比较成员的异同，不同则返回
                return false;

            if (this.Animation != srcFont.Animation) // 比较成员的异同，不同则返回
                return false;

            if (this.Bold != srcFont.Bold) // 比较成员的异同，不同则返回
                return false;

            if (this.BoldBi != srcFont.BoldBi) // 比较成员的异同，不同则返回
                return false;

            // copyBorders(srcFont.Borders, this.Borders); // 比较成员的异同，不同则返回
            // this.Borders.clone(srcFont.Borders);

            if (this.Color != srcFont.Color) // 比较成员的异同，不同则返回
                return false;

            if (this.ColorIndex != srcFont.ColorIndex) // 比较成员的异同，不同则返回
                return false;

            if (this.ColorIndexBi != srcFont.ColorIndexBi) // 比较成员的异同，不同则返回
                return false;

            if (this.DiacriticColor != srcFont.DiacriticColor) // 比较成员的异同，不同则返回
                return false;

            if (this.DisableCharacterSpaceGrid != srcFont.DisableCharacterSpaceGrid) // 比较成员的异同，不同则返回
                return false;

            if (this.DoubleStrikeThrough != srcFont.DoubleStrikeThrough) // 比较成员的异同，不同则返回
                return false;

            if (this.Emboss != srcFont.Emboss) // 比较成员的异同，不同则返回
                return false;

            if (this.EmphasisMark != srcFont.EmphasisMark) // 比较成员的异同，不同则返回
                return false;

            if (this.Engrave != srcFont.Engrave) // 比较成员的异同，不同则返回
                return false;

            if (this.Hidden != srcFont.Hidden) // 比较成员的异同，不同则返回
                return false;

            if (this.Italic != srcFont.Italic) // 比较成员的异同，不同则返回
                return false;

            if (this.ItalicBi != srcFont.ItalicBi) // 比较成员的异同，不同则返回
                return false;

            if (this.Kerning != srcFont.Kerning) // 比较成员的异同，不同则返回
                return false;

            if (this.Name != srcFont.Name) // 比较成员的异同，不同则返回
                return false;

            if (this.NameAscii != srcFont.NameAscii) // 比较成员的异同，不同则返回
                return false;

            if (this.NameBi != srcFont.NameBi) // 比较成员的异同，不同则返回
                return false;

            if (this.NameFarEast != srcFont.NameFarEast) // 比较成员的异同，不同则返回
                return false;

            //            if(this.NameOther != srcFont.NameOther) // 比较成员的异同，不同则返回
//                return false;

            if (this.Outline != srcFont.Outline) // 比较成员的异同，不同则返回
                return false;

            if (this.Position != srcFont.Position) // 比较成员的异同，不同则返回
                return false;

            if (this.Scaling != srcFont.Scaling) // 比较成员的异同，不同则返回
                return false;

            // this.Shading = srcFont.Shading;
            // this.Shading.clone(srcFont.Shading);

            if (this.Shadow != srcFont.Shadow) // 比较成员的异同，不同则返回
                return false;

            if (this.Size != srcFont.Size) // 比较成员的异同，不同则返回
                return false;

            if (this.SizeBi != srcFont.SizeBi) // 比较成员的异同，不同则返回
                return false;

            if (this.SmallCaps != srcFont.SmallCaps) // 比较成员的异同，不同则返回
                return false;

            if (this.Spacing != srcFont.Spacing) // 比较成员的异同，不同则返回
                return false;

            if (this.StrikeThrough != srcFont.StrikeThrough) // 比较成员的异同，不同则返回
                return false;

            if (this.Subscript != srcFont.Subscript) // 比较成员的异同，不同则返回
                return false;

            if (this.Superscript != srcFont.Superscript) // 比较成员的异同，不同则返回
                return false;

            if (this.Underline != srcFont.Underline) // 比较成员的异同，不同则返回
                return false;

            if (this.UnderlineColor != srcFont.UnderlineColor) // 比较成员的异同，不同则返回
                return false;

            return true;
        }


        // 复制WORD.FONT的内容到本类
        public void clone(Word.Font srcFont)
        {
            this.AllCaps = srcFont.AllCaps;// 复制WORD.FONT的成员到本类成员
            this.Animation = srcFont.Animation;// 复制WORD.FONT的成员到本类成员

            this.Bold = srcFont.Bold;// 复制WORD.FONT的成员到本类成员
            this.BoldBi = srcFont.BoldBi;// 复制WORD.FONT的成员到本类成员

            //this.Borders.clone(srcFont.Borders);

            this.Color = srcFont.Color;// 复制WORD.FONT的成员到本类成员
            this.ColorIndex = srcFont.ColorIndex;// 复制WORD.FONT的成员到本类成员
            this.ColorIndexBi = srcFont.ColorIndexBi;// 复制WORD.FONT的成员到本类成员

            this.DiacriticColor = srcFont.DiacriticColor;// 复制WORD.FONT的成员到本类成员
            this.DisableCharacterSpaceGrid = srcFont.DisableCharacterSpaceGrid;// 复制WORD.FONT的成员到本类成员
            this.DoubleStrikeThrough = srcFont.DoubleStrikeThrough;// 复制WORD.FONT的成员到本类成员

            this.Emboss = srcFont.Emboss;// 复制WORD.FONT的成员到本类成员
            this.EmphasisMark = srcFont.EmphasisMark;// 复制WORD.FONT的成员到本类成员
            this.Engrave = srcFont.Engrave;// 复制WORD.FONT的成员到本类成员
            this.Hidden = srcFont.Hidden;// 复制WORD.FONT的成员到本类成员
            this.Italic = srcFont.Italic;// 复制WORD.FONT的成员到本类成员
            this.ItalicBi = srcFont.ItalicBi;// 复制WORD.FONT的成员到本类成员
            this.Kerning = srcFont.Kerning;// 复制WORD.FONT的成员到本类成员
            
            this.NameAscii = srcFont.NameAscii;// 复制WORD.FONT的成员到本类成员
            this.NameBi = srcFont.NameBi;// 复制WORD.FONT的成员到本类成员
            this.NameFarEast = srcFont.NameFarEast;// 复制WORD.FONT的成员到本类成员
            //this.NameOther = srcFont.NameOther;// 复制WORD.FONT的成员到本类成员

            this.Name = srcFont.Name;// 复制WORD.FONT的成员到本类成员
            
            this.Outline = srcFont.Outline;// 复制WORD.FONT的成员到本类成员
            this.Position = srcFont.Position;// 复制WORD.FONT的成员到本类成员
            this.Scaling = srcFont.Scaling;// 复制WORD.FONT的成员到本类成员

            //this.Shading.clone(srcFont.Shading);// 复制WORD.FONT的成员到本类成员

            this.Shadow = srcFont.Shadow;// 复制WORD.FONT的成员到本类成员
            this.Size = srcFont.Size;// 复制WORD.FONT的成员到本类成员
            this.SizeBi = srcFont.SizeBi;// 复制WORD.FONT的成员到本类成员
            this.SmallCaps = srcFont.SmallCaps;// 复制WORD.FONT的成员到本类成员
            this.Spacing = srcFont.Spacing;// 复制WORD.FONT的成员到本类成员
            this.StrikeThrough = srcFont.StrikeThrough;// 复制WORD.FONT的成员到本类成员
            this.Subscript = srcFont.Subscript;// 复制WORD.FONT的成员到本类成员
            this.Superscript = srcFont.Superscript;// 复制WORD.FONT的成员到本类成员
            this.Underline = srcFont.Underline;// 复制WORD.FONT的成员到本类成员
            this.UnderlineColor = srcFont.UnderlineColor;// 复制WORD.FONT的成员到本类成员
           
            return;
        }


        // 复制本类到WORD.FONT类
        public void copy2(Word.Font dstFnt)
        {
            dstFnt.AllCaps = this.AllCaps;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Animation = this.Animation;// 复制本类成员到WORD.FONT类同名成员

            dstFnt.Bold = this.Bold;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.BoldBi = this.BoldBi;// 复制本类成员到WORD.FONT类同名成员

            Word.Borders bds = dstFnt.Borders;// 复制本类成员到WORD.FONT类同名成员
            //this.Borders.copy2(ref bds);

            dstFnt.Color = this.Color;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.ColorIndex = this.ColorIndex;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.ColorIndexBi = this.ColorIndexBi;// 复制本类成员到WORD.FONT类同名成员

            dstFnt.DiacriticColor = this.DiacriticColor;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.DisableCharacterSpaceGrid = this.DisableCharacterSpaceGrid;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.DoubleStrikeThrough = this.DoubleStrikeThrough;// 复制本类成员到WORD.FONT类同名成员

            dstFnt.Emboss = this.Emboss;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.EmphasisMark = this.EmphasisMark;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Engrave = this.Engrave;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Hidden = this.Hidden;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Italic = this.Italic;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.ItalicBi = this.ItalicBi;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Kerning = this.Kerning;// 复制本类成员到WORD.FONT类同名成员

            dstFnt.NameAscii = this.NameAscii;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.NameBi = this.NameBi;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.NameFarEast = this.NameFarEast;// 复制本类成员到WORD.FONT类同名成员
            //dstFnt.NameOther = this.NameOther;// 复制本类成员到WORD.FONT类同名成员
            
            dstFnt.Name = this.Name;// 复制本类成员到WORD.FONT类同名成员

            dstFnt.Outline = this.Outline;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Position = this.Position;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Scaling = this.Scaling;// 复制本类成员到WORD.FONT类同名成员

            Word.Shading shd = dstFnt.Shading;// 复制本类成员到WORD.FONT类同名成员
            //this.Shading.copy2(ref shd);

            dstFnt.Shadow = this.Shadow;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Size = this.Size;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.SizeBi = this.SizeBi;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.SmallCaps = this.SmallCaps;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Spacing = this.Spacing;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.StrikeThrough = this.StrikeThrough;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Subscript = this.Subscript;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Superscript = this.Superscript;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Underline = this.Underline;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.UnderlineColor = this.UnderlineColor;// 复制本类成员到WORD.FONT类同名成员
            
            return;
        }

        // 复制ClassFont到本类
        public void clone(ClassFont srcFont)
        {
            this.AllCaps = srcFont.AllCaps;// 复制ClassFont类同名成员到本类成员
            this.Animation = srcFont.Animation;// 复制ClassFont类同名成员到本类成员

            this.Bold = srcFont.Bold;// 复制ClassFont类同名成员到本类成员
            this.BoldBi = srcFont.BoldBi;// 复制ClassFont类同名成员到本类成员

            this.Borders.clone(srcFont.Borders);// 复制ClassFont类同名成员到本类成员

            this.Color = srcFont.Color;// 复制ClassFont类同名成员到本类成员
            this.ColorIndex = srcFont.ColorIndex;// 复制ClassFont类同名成员到本类成员
            this.ColorIndexBi = srcFont.ColorIndexBi;// 复制ClassFont类同名成员到本类成员

            this.DiacriticColor = srcFont.DiacriticColor;// 复制ClassFont类同名成员到本类成员
            this.DisableCharacterSpaceGrid = srcFont.DisableCharacterSpaceGrid;// 复制ClassFont类同名成员到本类成员
            this.DoubleStrikeThrough = srcFont.DoubleStrikeThrough;// 复制ClassFont类同名成员到本类成员

            this.Emboss = srcFont.Emboss;// 复制ClassFont类同名成员到本类成员
            this.EmphasisMark = srcFont.EmphasisMark;// 复制ClassFont类同名成员到本类成员
            this.Engrave = srcFont.Engrave;// 复制ClassFont类同名成员到本类成员
            this.Hidden = srcFont.Hidden;// 复制ClassFont类同名成员到本类成员
            this.Italic = srcFont.Italic;// 复制ClassFont类同名成员到本类成员
            this.ItalicBi = srcFont.ItalicBi;// 复制ClassFont类同名成员到本类成员
            this.Kerning = srcFont.Kerning;// 复制ClassFont类同名成员到本类成员
            
            this.NameAscii = srcFont.NameAscii;// 复制ClassFont类同名成员到本类成员
            this.NameBi = srcFont.NameBi;// 复制ClassFont类同名成员到本类成员
            this.NameFarEast = srcFont.NameFarEast;// 复制ClassFont类同名成员到本类成员
            //this.NameOther = srcFont.NameOther;// 复制ClassFont类同名成员到本类成员

            this.Name = srcFont.Name;// 复制ClassFont类同名成员到本类成员
            
            this.Outline = srcFont.Outline;// 复制ClassFont类同名成员到本类成员
            this.Position = srcFont.Position;// 复制ClassFont类同名成员到本类成员
            this.Scaling = srcFont.Scaling;// 复制ClassFont类同名成员到本类成员

            this.Shading.clone(srcFont.Shading);// 复制ClassFont类同名成员到本类成员

            this.Shadow = srcFont.Shadow;// 复制ClassFont类同名成员到本类成员
            this.Size = srcFont.Size;// 复制ClassFont类同名成员到本类成员
            this.SizeBi = srcFont.SizeBi;// 复制ClassFont类同名成员到本类成员
            this.SmallCaps = srcFont.SmallCaps;// 复制ClassFont类同名成员到本类成员
            this.Spacing = srcFont.Spacing;// 复制ClassFont类同名成员到本类成员
            this.StrikeThrough = srcFont.StrikeThrough;// 复制ClassFont类同名成员到本类成员
            this.Subscript = srcFont.Subscript;// 复制ClassFont类同名成员到本类成员
            this.Superscript = srcFont.Superscript;// 复制ClassFont类同名成员到本类成员
            this.Underline = srcFont.Underline;// 复制ClassFont类同名成员到本类成员
            this.UnderlineColor = srcFont.UnderlineColor;// 复制ClassFont类同名成员到本类成员

            return;
        }

        // 复制本类到ClassFont类
        public void copy2(ClassFont dstFnt)
        {
            dstFnt.AllCaps = this.AllCaps;// 复制本类成员到ClassFont类同名成员
            dstFnt.Animation = this.Animation;// 复制本类成员到ClassFont类同名成员

            dstFnt.Bold = this.Bold;// 复制本类成员到ClassFont类同名成员
            dstFnt.BoldBi = this.BoldBi;// 复制本类成员到ClassFont类同名成员

            ClassBorders bds = dstFnt.Borders;// 复制本类成员到ClassFont类同名成员
            this.Borders.copy2(ref bds);// 复制本类成员到ClassFont类同名成员

            dstFnt.Color = this.Color;// 复制本类成员到ClassFont类同名成员
            dstFnt.ColorIndex = this.ColorIndex;// 复制本类成员到ClassFont类同名成员
            dstFnt.ColorIndexBi = this.ColorIndexBi;// 复制本类成员到ClassFont类同名成员

            dstFnt.DiacriticColor = this.DiacriticColor;// 复制本类成员到ClassFont类同名成员
            dstFnt.DisableCharacterSpaceGrid = this.DisableCharacterSpaceGrid;// 复制本类成员到ClassFont类同名成员
            dstFnt.DoubleStrikeThrough = this.DoubleStrikeThrough;// 复制本类成员到ClassFont类同名成员

            dstFnt.Emboss = this.Emboss;// 复制本类成员到ClassFont类同名成员
            dstFnt.EmphasisMark = this.EmphasisMark;// 复制本类成员到ClassFont类同名成员
            dstFnt.Engrave = this.Engrave;// 复制本类成员到ClassFont类同名成员
            dstFnt.Hidden = this.Hidden;// 复制本类成员到ClassFont类同名成员
            dstFnt.Italic = this.Italic;// 复制本类成员到ClassFont类同名成员
            dstFnt.ItalicBi = this.ItalicBi;// 复制本类成员到ClassFont类同名成员
            dstFnt.Kerning = this.Kerning;// 复制本类成员到ClassFont类同名成员

            dstFnt.NameAscii = this.NameAscii;// 复制本类成员到ClassFont类同名成员
            dstFnt.NameBi = this.NameBi;// 复制本类成员到ClassFont类同名成员
            dstFnt.NameFarEast = this.NameFarEast;// 复制本类成员到ClassFont类同名成员
            //dstFnt.NameOther = this.NameOther;// 复制本类成员到ClassFont类同名成员

            dstFnt.Name = this.Name;// 复制本类成员到ClassFont类同名成员
            
            dstFnt.Outline = this.Outline;// 复制本类成员到ClassFont类同名成员
            dstFnt.Position = this.Position;// 复制本类成员到ClassFont类同名成员
            dstFnt.Scaling = this.Scaling;// 复制本类成员到ClassFont类同名成员

            ClassShading shd = dstFnt.Shading;// 复制本类成员到ClassFont类同名成员
            this.Shading.copy2(ref shd);// 复制本类成员到ClassFont类同名成员

            dstFnt.Shadow = this.Shadow;// 复制本类成员到ClassFont类同名成员
            dstFnt.Size = this.Size;// 复制本类成员到ClassFont类同名成员
            dstFnt.SizeBi = this.SizeBi;// 复制本类成员到ClassFont类同名成员
            dstFnt.SmallCaps = this.SmallCaps;// 复制本类成员到ClassFont类同名成员
            dstFnt.Spacing = this.Spacing;// 复制本类成员到ClassFont类同名成员
            dstFnt.StrikeThrough = this.StrikeThrough;// 复制本类成员到ClassFont类同名成员
            dstFnt.Subscript = this.Subscript;// 复制本类成员到ClassFont类同名成员
            dstFnt.Superscript = this.Superscript;// 复制本类成员到ClassFont类同名成员
            dstFnt.Underline = this.Underline;// 复制本类成员到ClassFont类同名成员
            dstFnt.UnderlineColor = this.UnderlineColor;// 复制本类成员到ClassFont类同名成员

            return;
        }


    }
}
