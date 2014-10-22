using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;


namespace CppNav
{
  [Export(typeof(IVsTextViewCreationListener))]
  [ContentType("text")]
  [TextViewRole(PredefinedTextViewRoles.Editable)]
  internal class KeyBindingFilterProvider : IVsTextViewCreationListener
  {
    public void VsTextViewCreated(IVsTextView text_view_adapter)
    {
      IWpfTextView text_view = editor_factory.GetWpfTextView(text_view_adapter);
      if (text_view == null) return;
      AddCommandFilter(text_view_adapter, new KeyBindingCommandFilter(text_view));
    }

    void AddCommandFilter(IVsTextView view_adapter, KeyBindingCommandFilter cmd_filter)
    {
      if (cmd_filter.is_added == false) {
        IOleCommandTarget next;
        int hres = view_adapter.AddCommandFilter(cmd_filter, out next);

        if (hres == VSConstants.S_OK) {
          cmd_filter.is_added = true;
          if (next != null) cmd_filter.next_target = next;
        }
      }
    }

    [Import(typeof(IVsEditorAdaptersFactoryService))]
    internal IVsEditorAdaptersFactoryService editor_factory = null;
  }
}
