using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Collections.Generic;

namespace VisualStudio.SubwordNavigation {
    internal class SubwordNavigationCommandFilter : IOleCommandTarget {
        private readonly IWpfTextView textView;
        private readonly IEditorOperations operations;
        private readonly ITextStructureNavigator navigator;
        private readonly Dictionary<uint, Action> commands = new Dictionary<uint, Action>();

        public SubwordNavigationCommandFilter(IWpfTextView textView, IEditorOperations operations, ITextStructureNavigator navigator) {
            this.textView = textView;
            this.operations = operations;
            this.navigator = navigator;

            RegisterCommands();
        }

        internal IOleCommandTarget Next { get; set; }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut) {
            if (pguidCmdGroup != GuidList.guidSubwordNavigationCmdSet) {
                return Next.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }

            commands[nCmdID]?.Invoke();

            return VSConstants.S_OK;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText) {
            if (pguidCmdGroup != GuidList.guidSubwordNavigationCmdSet) {
                return Next.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
            }

            for (int i = 0; i < prgCmds.Length; i++) {
                if (commands.ContainsKey(prgCmds[i].cmdID)) {
                    prgCmds[i].cmdf = (uint)(OLECMDF.OLECMDF_ENABLED | OLECMDF.OLECMDF_SUPPORTED);
                }
            }

            return VSConstants.S_OK;
        }

        private void RegisterCommands() {
            commands.Add(PkgCmdIDList.cmdidSubwordPrevious, () => MovePrevious(false));
            commands.Add(PkgCmdIDList.cmdidSubwordNext, () => MoveNext(false));

            commands.Add(PkgCmdIDList.cmdidSubwordPreviousExtend, () => MovePrevious(true));
            commands.Add(PkgCmdIDList.cmdidSubwordNextExtend, () => MoveNext(true));

            commands.Add(PkgCmdIDList.cmdidSubwordDeletePrevious, () => Delete(MovePrevious));
            commands.Add(PkgCmdIDList.cmdidSubwordDeleteNext, () => Delete(MoveNext));

            commands.Add(PkgCmdIDList.cmdidSubwordShrinkSelection, () => ShrinkSelection());
            commands.Add(PkgCmdIDList.cmdidSubwordGrowSelection, () => GrowSelection());
        }

        private void MovePrevious(bool extendSelection) {
            if (!extendSelection && !textView.Selection.IsEmpty) {
                textView.Selection.Select(textView.Selection.ActivePoint, textView.Selection.ActivePoint);
            }

            var caret = textView.Caret;

            if (caret.InVirtualSpace) {
                operations.MoveToEndOfLine(extendSelection);
                return;
            }

            var point = caret.Position.BufferPosition;

            if (point.Position == 0) {
                return;
            }

            if (point == caret.ContainingTextViewLine.Start) {
                operations.MoveLineUp(extendSelection);
                operations.MoveToLastNonWhiteSpaceCharacter(extendSelection);
                if (navigator.GetExtentOfWord(caret.Position.BufferPosition).IsSignificant) {
                    operations.MoveToNextCharacter(extendSelection);
                }
                return;
            }

            var extent = navigator.GetExtentOfSubword(point - 1);

            if (extent == null) {
                return;
            }

            for (int i = point; i > extent.Value.Span.Start; i--) {
                operations.MoveToPreviousCharacter(extendSelection);
            }
        }

        private void MoveNext(bool extendSelection) {
            if (!extendSelection && !textView.Selection.IsEmpty) {
                textView.Selection.Select(textView.Selection.ActivePoint, textView.Selection.ActivePoint);
            }

            var caret = textView.Caret;
            var point = caret.Position.BufferPosition;

            if (point.Position == textView.TextSnapshot.Length) {
                return;
            }

            if (point == caret.ContainingTextViewLine.End) {
                operations.MoveToStartOfNextLineAfterWhiteSpace(extendSelection);  
                return;
            }

            var extent = navigator.GetExtentOfSubword(point);

            if (extent == null) {
                return;
            }

            for (int i = point; i < extent.Value.Span.End && i < caret.ContainingTextViewLine.End; i++) {
                operations.MoveToNextCharacter(extendSelection);
            }
        }

        private void Delete(Action<bool> select) {
            if (textView.Selection.IsEmpty) {
                select(true);
            }
            operations.Delete();
        }

        private void ShrinkSelection()
        {
            var caret = textView.Caret;

            if (caret.InVirtualSpace)
            {
                return;
            }

            var point = caret.Position.BufferPosition;

            if (point.Position == 0)
            {
                return;
            }

            if (textView.Selection.IsEmpty)
            {
                return;
            }

            var startExtent = navigator.GetExtentOfSubword(textView.Selection.Start.Position);
            var endExtent = navigator.GetExtentOfSubword(textView.Selection.End.Position - 1);

            if (startExtent == null || endExtent == null)
            {
                return;
            }

            if (startExtent.Value.Span == endExtent.Value.Span)
            {
                textView.Selection.Select(textView.Selection.ActivePoint, textView.Selection.ActivePoint);
                return;
            }

            // Prefer shrinking other characters or words to shrinking the subword
            char leftChar = startExtent.Value.Span.End.GetChar();
            char rightChar = endExtent.Value.Span.Start.Subtract(1).GetChar();
            var preferLeft = char.IsLetterOrDigit(leftChar) || leftChar == '_';
            var preferRight = char.IsLetterOrDigit(rightChar) || rightChar == '_';

            // Compare against caret, prefer to keep the subword under the caret selected
            preferLeft &= startExtent.Value.Span.End >= point;
            preferRight &= endExtent.Value.Span.Start <= point;

            if ((preferLeft && preferRight) || (!preferLeft && !preferRight))
            {
                textView.Selection.Select(new SnapshotSpan(startExtent.Value.Span.End, endExtent.Value.Span.Start), false);
            }
            else if (preferLeft)
            {
                textView.Selection.Select(new SnapshotSpan(textView.Selection.Start.Position, endExtent.Value.Span.Start), false);
            }
            else
            {
                textView.Selection.Select(new SnapshotSpan(startExtent.Value.Span.End, textView.Selection.End.Position), false);
            }
        }
        private void GrowSelection()
        {
            var caret = textView.Caret;

            if (caret.InVirtualSpace)
            {
                return;
            }

            var point = caret.Position.BufferPosition;

            if (point.Position == 0)
            {
                return;
            }

            var startPoint = textView.Selection.IsEmpty ? point : textView.Selection.Start.Position;
            var endPoint = textView.Selection.IsEmpty ? point : textView.Selection.End.Position;
            var startExtent = navigator.GetExtentOfSubword(startPoint - 1);
            var endExtent = navigator.GetExtentOfSubword(endPoint);

            if (startExtent == null || endExtent == null)
            {
                return;
            }

            // Prefer extending the subword over extending to include other characters or words
            char leftChar = startExtent.Value.Span.End.Subtract(1).GetChar();
            char rightChar = endExtent.Value.Span.Start.GetChar();
            var isLeftLetterOrDigit = char.IsLetterOrDigit(leftChar) || leftChar == '_';
            var isRightLetterOrDigit = char.IsLetterOrDigit(rightChar) || rightChar == '_';

            if ((isLeftLetterOrDigit && isRightLetterOrDigit) || (!isLeftLetterOrDigit && !isRightLetterOrDigit))
            {
                textView.Selection.Select(new SnapshotSpan(startExtent.Value.Span.Start, endExtent.Value.Span.End), false);
            }
            else if (isLeftLetterOrDigit)
            {
                textView.Selection.Select(new SnapshotSpan(startExtent.Value.Span.Start, endPoint), false);
            }
            else
            {
                textView.Selection.Select(new SnapshotSpan(startPoint, endExtent.Value.Span.End), false);
            }
        }
    }
}
