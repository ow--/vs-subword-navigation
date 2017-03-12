using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;
using System.Diagnostics;

namespace VisualStudio.SubwordNavigation {
    [Export(typeof(IVsTextViewCreationListener))]
    [ContentType("text")]
    [TextViewRole(PredefinedTextViewRoles.Interactive)]
    internal class SubwordNavigationCommandFilterProvider : IVsTextViewCreationListener {
        [Import]
        internal IVsEditorAdaptersFactoryService editorAdaptersFactory = null;
        [Import]
        internal IEditorOperationsFactoryService editorOperationsFactory = null;
        [Import]
        internal ITextStructureNavigatorSelectorService textStructureNavigatorFacrtory = null;

        public void VsTextViewCreated(IVsTextView textViewAdapter) {
            var textView = editorAdaptersFactory.GetWpfTextView(textViewAdapter);
            var operations = editorOperationsFactory.GetEditorOperations(textView);
            var navigator = textStructureNavigatorFacrtory.GetTextStructureNavigator(textView.TextBuffer);

            if (textView == null) {
                Debug.Fail("Unable to get IWpfTextView from text view adapter.");
                return;
            }

            var commandFilter = new SubwordNavigationCommandFilter(textView, operations, navigator);

            IOleCommandTarget next;
            var hr = textViewAdapter.AddCommandFilter(commandFilter, out next);
            if (ErrorHandler.Succeeded(hr)) {
                commandFilter.Next = next;
            }
        }
    }
}
