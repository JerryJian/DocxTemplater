using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;

namespace DocxTemplater.Markdown
{
#if NET6_0_OR_GREATER
    public record ListLevelConfiguration(
        string LevelText,
        string FontOverride,
        NumberFormatValues NumberingFormat,
        int IndentPerLevel);
#else
    public class ListLevelConfiguration : IEquatable<ListLevelConfiguration>
    {
        public ListLevelConfiguration(string levelText, string fontOverride, NumberFormatValues numberingFormat, int indentPerLevel)
        {
            LevelText = levelText;
            FontOverride = fontOverride;
            NumberingFormat = numberingFormat;
            IndentPerLevel = indentPerLevel;
        }

        public string LevelText { get; }

        public string FontOverride { get; }

        public NumberFormatValues NumberingFormat { get; }

        public int IndentPerLevel { get; }

        public override bool Equals(object obj)
        {
            return Equals(obj as ListLevelConfiguration);
        }

        public bool Equals(ListLevelConfiguration other)
        {
            return other is not null &&
                   LevelText == other.LevelText &&
                   FontOverride == other.FontOverride &&
                   NumberingFormat.Equals(other.NumberingFormat) &&
                   IndentPerLevel == other.IndentPerLevel;
        }

        public override int GetHashCode()
        {
            var hashCode = 1103593537;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(LevelText);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(FontOverride);
            hashCode = hashCode * -1521134295 + NumberingFormat.GetHashCode();
            hashCode = hashCode * -1521134295 + IndentPerLevel.GetHashCode();
            return hashCode;
        }
    }
#endif

    public class MarkDownFormatterConfiguration
    {
        public static readonly MarkDownFormatterConfiguration Default;

        static MarkDownFormatterConfiguration()
        {
            Default = new MarkDownFormatterConfiguration();
        }

        public MarkDownFormatterConfiguration()
        {
            OrderedListLevelConfiguration = new List<ListLevelConfiguration>();
            UnorderedListLevelConfiguration = new List<ListLevelConfiguration>();

            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%1.", null, NumberFormatValues.Decimal, 720));
            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%2.", null, NumberFormatValues.LowerLetter, 720));
            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%3.", null, NumberFormatValues.LowerRoman, 720));

            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%4.", null, NumberFormatValues.Decimal, 720));
            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%5.", null, NumberFormatValues.LowerLetter, 720));
            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%6.", null, NumberFormatValues.LowerRoman, 720));

            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%7.", null, NumberFormatValues.Decimal, 720));
            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%8.", null, NumberFormatValues.LowerLetter, 720));
            OrderedListLevelConfiguration.Add(new ListLevelConfiguration("%9.", null, NumberFormatValues.LowerRoman, 720));

            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("\uf0b7", "Symbol", NumberFormatValues.Bullet, 720));
            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("o", "Courier New", NumberFormatValues.Bullet, 720));
            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("\uf0a7", "Wingdings", NumberFormatValues.Bullet, 720));

            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("\uf0b7", "Symbol", NumberFormatValues.Bullet, 720));
            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("o", "Courier New", NumberFormatValues.Bullet, 720));
            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("\uf0a7", "Wingdings", NumberFormatValues.Bullet, 720));

            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("\uf0b7", "Symbol", NumberFormatValues.Bullet, 720));
            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("o", "Courier New", NumberFormatValues.Bullet, 720));
            UnorderedListLevelConfiguration.Add(new ListLevelConfiguration("\uf0a7", "Wingdings", NumberFormatValues.Bullet, 720));
        }

        public List<ListLevelConfiguration> OrderedListLevelConfiguration
        {
            get;
        }

        public List<ListLevelConfiguration> UnorderedListLevelConfiguration
        {
            get;
        }

        /// <summary>
        /// Name of a list style in the template document applied to lists.
        /// If this style is not found, a style is created based on <see cref="UnorderedListLevelConfiguration"/>
        /// </summary>
        public string UnorderedListStyle { get; set; } = "md_ListStyle";

        /// <summary>
        /// Name of a list style in the template document applied to lists.
        /// If this style is not found, a style is created based on <see cref="OrderedListLevelConfiguration"/>
        /// </summary>
        public string OrderedListStyle { get; set; } = "md_OrderedListStyle";


        /// <summary>
        /// Name of a table style in the template document applied to tables.
        /// </summary>
        public string TableStyle { get; set; } = "md_TableStyle";
    }
}
