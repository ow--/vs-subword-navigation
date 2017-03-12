using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace VisualStudio.SubwordNavigation {
    public static class ITextStructureNavigatorExtensions {
        public static TextExtent? GetExtentOfSubword(this ITextStructureNavigator navigator, SnapshotPoint currentPosition) {
            var wordExtent = navigator.GetExtentOfWord(currentPosition);
            if (wordExtent.Span.Length == 0) {
                return null;
            }
            var word = wordExtent.Span.GetText();
            var subwords = Regex.Split(word, @"(_+)|(?=\p{Lu}\p{Ll})|(?<=\p{Ll})(?=\p{Lu})", RegexOptions.Compiled);
            var extents = new List<TextExtent>();
            var start = wordExtent.Span.Start;
            foreach (var item in subwords) {
                if (currentPosition.Position >= start && currentPosition.Position < start + item.Length) {
                    return new TextExtent(new SnapshotSpan(wordExtent.Span.Snapshot, start, item.Length), true);
                }
                start += item.Length;
            }
            return null;
        }
    }
}
