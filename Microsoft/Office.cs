using System;
using System.Collections.Generic;

namespace Microsoft
{
    internal class Office
    {
        internal class Interop
        {
            internal class Word
            {
                internal class Application
                {
                    public object Documents { get; internal set; }
                }

                internal class Document
                {
                    public object ActiveWindow { get; internal set; }
                    public IEnumerable<Paragraph> Paragraphs { get; internal set; }

                    public static implicit operator iText.Layout.Document(Document v)
                    {
                        throw new NotImplementedException();
                    }

                    internal void Close(ref object nullobject1, ref object nullobject2, ref object nullobject3)
                    {
                        throw new NotImplementedException();
                    }
                }

                internal class Paragraph
                {
                }

                internal class Font
                {
                    internal readonly object ColorIndex;
                    internal object Name;
                    internal object Position;
                    internal float Size;
                }
            }
        }
    }
}