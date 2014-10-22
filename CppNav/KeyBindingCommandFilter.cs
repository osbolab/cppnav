using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;


namespace CppNav
{
  internal class KeyBindingCommandFilter : IOleCommandTarget
  {
    [DllImport("user32.dll", SetLastError = true)]
    static extern byte GetKeyState(VirtualKey virtual_key);

    public enum VirtualKey : int
    {
      Shift = 0x10
    }

    private static bool is_key(VirtualKey v_key)
    {
      switch (GetKeyState(v_key)) {
        case 0:
        case 1: return false;
        default: return true;
      }
    }

    private byte[] keystate = new byte[256];

    public KeyBindingCommandFilter(IWpfTextView text_view)
    {
      this.text_view = text_view;
    }

    int IOleCommandTarget.QueryStatus(ref Guid group_id,
                                      uint nr_cmds, OLECMD[] cmds,
                                      IntPtr text)
    {
      return next_target.QueryStatus(group_id, nr_cmds, cmds, text);
    }

    int IOleCommandTarget.Exec(ref Guid group_id, uint id, uint exec_opt,
                               IntPtr va_in, IntPtr va_out)
    {
      if (group_id == VSConstants.VSStd2K) {
        if (id == (uint) VSConstants.VSStd2KCmdID.DELETEWORDLEFT) {
          deletePrevToken();
          return VSConstants.S_OK;
        } else if (id == (uint) VSConstants.VSStd2KCmdID.DELETEWORDRIGHT) {
          deleteNextToken();
         return VSConstants.S_OK;
        } else if (id == (uint) VSConstants.VSStd2KCmdID.WORDPREV) {
          movePrevToken(false);
          return VSConstants.S_OK;
        } else if (id == (uint) VSConstants.VSStd2KCmdID.WORDPREV_EXT) {
          movePrevToken(true);
          return VSConstants.S_OK;
        } else if (id == (uint) VSConstants.VSStd2KCmdID.WORDNEXT) {
          moveNextToken(false);
          return VSConstants.S_OK;
        } else if (id == (uint) VSConstants.VSStd2KCmdID.WORDNEXT_EXT) {
          moveNextToken(true);
          return VSConstants.S_OK;
        }
      }
      return next_target.Exec(ref group_id, id, exec_opt, va_in, va_out);
    }

    private void movePrevToken(bool select)
    {
      int caret = text_view.Caret.Position.BufferPosition.Position;
      if (caret <= 0) return;

      ITextSnapshot snapshot = text_view.TextSnapshot;
      if (snapshot != snapshot.TextBuffer.CurrentSnapshot) return;

      ITextSelection sel = text_view.Selection;
      int anchor = -1;
      if (select) {
        if (sel.SelectedSpans.Count > 0) {
          anchor = sel.AnchorPoint.Position.Position;
        }
      }

      Span token = new CppTokenizer(snapshot, caret, CppTokenizer.Direction.Back).next();
      text_view.Caret.MoveTo(new SnapshotPoint(snapshot, token.Start));

      if (select) {
        if (anchor < 0) anchor = token.End;
        sel.Select(new SnapshotSpan(snapshot, Span.FromBounds(token.Start, anchor)), (anchor > token.Start));
      } else {
        sel.Clear();
      }
    }

    private void moveNextToken(bool select)
    {
      int caret = text_view.Caret.Position.BufferPosition.Position;

      ITextSnapshot snapshot = text_view.TextSnapshot;
      if (snapshot != snapshot.TextBuffer.CurrentSnapshot) return;
      if (caret >= snapshot.Length) return;

      ITextSelection sel = text_view.Selection;
      int anchor = -1;
      if (select) {
        if (sel.SelectedSpans.Count > 0) {
          anchor = sel.AnchorPoint.Position.Position;
        }
      }

      Span token = new CppTokenizer(snapshot, caret, CppTokenizer.Direction.Forward).next();
      text_view.Caret.MoveTo(new SnapshotPoint(snapshot, token.End));
      
      if (select) {
        if (anchor < 0) anchor = token.Start;
        sel.Select(new SnapshotSpan(snapshot, Span.FromBounds(anchor, token.End)), (anchor > token.End));
      } else {
        sel.Clear();
      }
    }

    private void deletePrevToken()
    {
      int caret = text_view.Caret.Position.BufferPosition.Position;
      if (caret <= 0) return;

      ITextSnapshot snapshot = text_view.TextSnapshot;
      if (snapshot != snapshot.TextBuffer.CurrentSnapshot) return;

      // Start from the character behind the caret
      delete_token(snapshot, new CppTokenizer(snapshot, caret, CppTokenizer.Direction.Back));
    }

    private void deleteNextToken()
    {
      int caret = text_view.Caret.Position.BufferPosition.Position;

      ITextSnapshot snapshot = text_view.TextSnapshot;
      if (snapshot != snapshot.TextBuffer.CurrentSnapshot) return;
      if (caret >= snapshot.Length) return;

      delete_token(snapshot, new CppTokenizer(snapshot, caret, CppTokenizer.Direction.Forward));
    }

    private void delete_token(ITextSnapshot snapshot, CppTokenizer tokenizer)
    {
      Span token = tokenizer.next();
      using (var edit = snapshot.TextBuffer.CreateEdit()) {
        edit.Replace(token, "");
        edit.Apply();
      }
    }

    private IWpfTextView text_view;
    internal IOleCommandTarget next_target;
    internal bool is_added;
  }

  enum TokenType { Symbol, White, Delim };

  internal struct CppToken
  {
    internal CppToken(TokenType type)
    {
      this.type = type;
      start = -1;
      end = -1;
    }
    internal TokenType type;
    internal int start;
    internal int end;
  }

  class CppTokenizer
  {
    internal enum Direction { Back = -1, Forward = 1 };

    internal CppTokenizer(ITextSnapshot snapshot, int caret, Direction dir)
    {
      this.snapshot = snapshot;
      this.dir = dir;
      this.caret = (dir == Direction.Back) ? caret - 1 : caret;
    }

    internal Span next()
    {
      TokenType type = classify(snapshot[caret]);
      CppToken token;

      switch (type) {
        case TokenType.Symbol: token = next_symbol(); break;
        case TokenType.Delim: token = next_delim(1); break;
        case TokenType.White:
        default: token = next_white(); break;
      }

      return Span.FromBounds(token.start, token.end);
    }

    private CppToken next_symbol()
    {
      var token = start_token(TokenType.Symbol);

      char last = snapshot[caret];

      while (!is_eof()) {
        char c = peek();
        if (!is_symbol(c)) break;
        if (word_boundary(last, c)) break;
        last = c;
        take();
      }

      return end_token(token);
    }

    private CppToken next_delim(int count)
    {
      var token = start_token(TokenType.Delim);

      char last = snapshot[caret];

      while (!is_eof()) {
        char c = peek();
        if (c != last && !compound_delim(last, c)) {
          // Special case "for_" <-- that where deleting the _ deletes the word
          if (count == 1) {
            if (is_symbol(c)) {
              last = (char) 0; take(); next_symbol(); break;
            } 
            if (is_white(c)) {
              last = (char) 0; take(); next_white(); break;
            }
          }
          break;
        }
        ++count;
        last = c;
        take();
      }

      return end_token(token);
    }

    private CppToken next_white()
    {
      var token = start_token(TokenType.White);

      int count = 0;
      bool white_delim = false;
      while (!is_eof()) {
        ++count;
        char c = peek();
        if (!is_white(c)) {
          if (!white_delim && white_delims.Contains(c)) {
            take(); next_delim(2); white_delim = true; continue;
          }
          // Special case for one whitespace followed by something else
          if (count == 1) {
            if (is_symbol(c)) {
              take(); next_symbol(); break;
            } 
            if (is_delim(c)) {
              take(); next_delim(2); white_delim = true; break;
            }
          }
          break;
        }
        take();
      }

      return end_token(token);
    }

    private static char[] operators = { '+', '-', '*', '/' };
    private static char[] white_delims = { '+', '-', '*', '/', '=', '(', ')', ',' };

    private TokenType classify(char c)
    {
      if (is_white(c)) return TokenType.White;
      if (is_symbol(c)) return TokenType.Symbol;
      return TokenType.Delim;
    }

    private CppToken start_token(TokenType type)
    {
      var token = new CppToken(type);
      if (dir == Direction.Back) token.end = caret + 1;
      else token.start = caret;
      return token;
    }

    private CppToken end_token(CppToken token)
    {
      if (dir == Direction.Back) token.start = caret;
      else token.end = caret + 1;
      return token;

    }

    private bool is_eof()
    {
      if (dir == Direction.Back) return caret <= 0;
      else return caret >= snapshot.Length-1;
    }

    private char take()
    {
      if (dir == Direction.Back) return snapshot[--caret];
      return snapshot[++caret];
    }

    private char peek()
    {
      if (dir == Direction.Back) return snapshot[caret - 1];
      return snapshot[caret + 1];
    }

    private bool compound_delim(char last, char c)
    {
      if (last == '=') return operators.Contains(c);
      if (c == '=') return operators.Contains(last);
      return false;
    }

    private bool word_boundary(char last, char c)
    {
      if (dir == Direction.Back) return Char.IsUpper(last) && Char.IsLower(c);
      return Char.IsLower(last) && Char.IsUpper(c);
    }

    private static bool is_white(char c)
    {
      return Char.IsWhiteSpace(c);
    }

    private static bool is_symbol(char c)
    {
      return Char.IsLetterOrDigit(c);
    }

    private static bool is_delim(char c)
    {
      return !is_white(c) && !is_symbol(c);
    }

    private readonly ITextSnapshot snapshot;
    private int caret;
    private Direction dir;
  }
}
