using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using System;

namespace VisualStudio.SubwordNavigation {
    internal class SubwordNavigationCommandFilter : IOleCommandTarget {
        IWpfTextView textView;
        IEditorOperations operations;
        ITextStructureNavigator navigator;

        public SubwordNavigationCommandFilter(IWpfTextView textView, IEditorOperations operations, ITextStructureNavigator navigator) {
            this.textView = textView;
            this.operations = operations;
            this.navigator = navigator;
        }

        internal IOleCommandTarget Next { get; set; }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut) {
            if (pguidCmdGroup == GuidList.guidSubwordNavigationCmdSet) {
                if (nCmdID == PkgCmdIDList.cmdidSubwordPrevious) {
                    MovePrevious(false);
                } else if (nCmdID == PkgCmdIDList.cmdidSubwordNext) {
                    MoveNext(false);
                } else if (nCmdID == PkgCmdIDList.cmdidSubwordPreviousExtend) {
                    MovePrevious(true);
                } else if (nCmdID == PkgCmdIDList.cmdidSubwordNextExtend) {
                    MoveNext(true);
                } else if (nCmdID == PkgCmdIDList.cmdidSubwordDeletePrevious) {
                    if (textView.Selection.IsEmpty) {
                        MovePrevious(true);
                    }
                    operations.Delete();
                } else if (nCmdID == PkgCmdIDList.cmdidSubwordDeleteNext) {
                    if (textView.Selection.IsEmpty) {
                        MoveNext(true);
                    }
                    operations.Delete();
                }
                return VSConstants.S_OK;
            }
            return Next.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText) {
            if (pguidCmdGroup == GuidList.guidSubwordNavigationCmdSet) {
                for (int i = 0; i < prgCmds.Length; i++) {
                    if (prgCmds[i].cmdID == PkgCmdIDList.cmdidSubwordPrevious ||
                        prgCmds[i].cmdID == PkgCmdIDList.cmdidSubwordNext ||
                        prgCmds[i].cmdID == PkgCmdIDList.cmdidSubwordPreviousExtend ||
                        prgCmds[i].cmdID == PkgCmdIDList.cmdidSubwordNextExtend ||
                        prgCmds[i].cmdID == PkgCmdIDList.cmdidSubwordDeletePrevious ||
                        prgCmds[i].cmdID == PkgCmdIDList.cmdidSubwordDeleteNext
                        ) {
                        prgCmds[i].cmdf = (uint)(OLECMDF.OLECMDF_ENABLED | OLECMDF.OLECMDF_SUPPORTED);
                    }
                }
                return VSConstants.S_OK;
            }
            return Next.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        private void MovePrevious(bool extendSelection) {
            if (!extendSelection && !textView.Selection.IsEmpty) {
                textView.Selection.Select(textView.Selection.ActivePoint, textView.Selection.ActivePoint);
            }
            var point = textView.Caret.Position.BufferPosition;
            if (point.Position == 0) {
                return;
            }
            var extent = navigator.GetExtentOfSubword(point - 1);
            if (extent == null) {
                return;
            }
            for (int i = point.Position; i > extent.Value.Span.Start.Position; i--) {
                operations.MoveToPreviousCharacter(extendSelection);
            }
        }

        private void MoveNext(bool extendSelection) {
            if (!extendSelection && !textView.Selection.IsEmpty) {
                textView.Selection.Select(textView.Selection.ActivePoint, textView.Selection.ActivePoint);
            }
            var point = textView.Caret.Position.BufferPosition;
            var extent = navigator.GetExtentOfSubword(point);
            if (extent == null) {
                return;
            }
            for (int i = point.Position; i < extent.Value.Span.End.Position; i++) {
                operations.MoveToNextCharacter(extendSelection);
            }
        }
    }
}
