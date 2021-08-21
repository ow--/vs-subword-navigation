using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using System.Text.RegularExpressions;

namespace VisualStudio.SubwordNavigation {
    public static class ITextStructureNavigatorExtensions {
        public static TextExtent? GetExtentOfSubword(this ITextStructureNavigator navigator, SnapshotPoint currentPosition, ITextSnapshot textSnapshot) {
            var span = navigator.GetExtentOfWord(currentPosition).Span;
            // Depending on the type of file, sometimes the extent of the word won't exist,
            // or it will be the span of the word that the current position ends at rather than
            // the word that starts at that position (`.bat` files are an example where this occurs).
            if (span.IsEmpty || (span.End == currentPosition)) {
                // Get the span of the next element. We can use
                // this to work out where the next word really is.
                var next = navigator.GetSpanOfNextSibling(span);
                if (next.IsEmpty) {
                    return null;
                }
                if (next.Contains(currentPosition)) {
                    // The next element contains the current position (usually the current position is
                    // at the start, but for a whitespace span the current position could also be in the
                    // middle of it). The next element is the span we can use. An example of this is
                    // a .bat file that contains "test!" where the current position is in between
                    // the "t" and "!". The extent of the word is "test", and the next sibling is "!".
                    span = next;
                
                } else {
                    // The next element starts after the current position, so the span that we need to
                    // use is the gap between the current position and the start of the next sibling.
                    // An example of this is a .bat file that contains "test span" where the current
                    // position is at the end of "test". The extent of the word is "test", the next sibling
                    // is "span", and the span that we want to use is the space between those words.
                    span = new SnapshotSpan(textSnapshot, new Span(currentPosition, next.Start - currentPosition));
                }
            }
            var word = span.GetText();
            var subwords = Regex.Split(word, @"(_+)|(?=\p{Lu}\p{Ll})|(?<=\p{Ll})(?=\p{Lu})", RegexOptions.Compiled);
            var start = span.Start;
            foreach (var item in subwords) {
                if (currentPosition.Position >= start && currentPosition.Position < start + item.Length) {
                    return new TextExtent(new SnapshotSpan(span.Snapshot, start, item.Length), true);
                }
                start += item.Length;
            }
            return null;
        }
    }
}
