﻿using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformShading : PlatformOpenXmlElement, IShading
    {
        private readonly Shading xmlElement;

        public PlatformShading(Shading shading)
            : base(shading)
        {
            this.xmlElement = shading;
        }

        #region Interface

        public string Color
        {
            get
            {
                return xmlElement.Color;
            }

            set
            {
                xmlElement.Color = value;
            }
        }

        public string Fill
        {
            get
            {
                return xmlElement.Fill;
            }

            set
            {
                xmlElement.Fill = value;
            }
        }

        public Core.Word.ThemeColorValues? ThemeColor
        {
            get
            {
                if (xmlElement.ThemeColor != null && xmlElement.ThemeColor.HasValue)
                    return (Core.Word.ThemeColorValues)(int)xmlElement.ThemeColor.Value;
                else
                    return null;
            }

            set
            {
                if (value.HasValue)
                    xmlElement.ThemeColor = (DocumentFormat.OpenXml.Wordprocessing.ThemeColorValues)(int)value;
                else
                    xmlElement.ThemeColor = null;
            }
        }

        public Core.Word.ThemeColorValues? ThemeFill
        {
            get
            {
                if (xmlElement.ThemeFill != null && xmlElement.ThemeFill.HasValue)
                    return (Core.Word.ThemeColorValues)(int)xmlElement.ThemeFill.Value;
                else
                    return null;
            }

            set
            {
                if (value.HasValue)
                    xmlElement.ThemeFill = (DocumentFormat.OpenXml.Wordprocessing.ThemeColorValues)(int)value;
                else
                    xmlElement.ThemeFill = null;
            }
        }

        public string ThemeFillShade
        {
            get
            {
                return xmlElement.ThemeFillShade;
            }

            set
            {
                xmlElement.ThemeFillShade = value;
            }
        }

        public string ThemeFillTint
        {
            get
            {
                return xmlElement.ThemeFillTint;
            }

            set
            {
                xmlElement.ThemeFillTint = value;
            }
        }

        public string ThemeShade
        {
            get
            {
                return xmlElement.ThemeShade;
            }

            set
            {
                xmlElement.ThemeShade = value;
            }
        }

        public string ThemeTint
        {
            get
            {
                return xmlElement.ThemeTint;
            }

            set
            {
                xmlElement.ThemeTint = value;
            }
        }

        public Core.Word.ShadingPatternValues? Val
        {
            get
            {
                if (xmlElement.Val != null && xmlElement.Val.HasValue)
                    return (Core.Word.ShadingPatternValues)(int)xmlElement.Val.Value;
                else
                    return null;
            }

            set
            {
                if (value.HasValue)
                    xmlElement.Val = (DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues)(int)value;
                else
                    xmlElement.Val = null;
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformShading New(TableProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<Shading>(tableProperties);
            return new PlatformShading(xmlElement);
        }

        public static PlatformShading New(TableCellProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<Shading>(tableProperties);
            return new PlatformShading(xmlElement);
        }

        #endregion
    }
}
