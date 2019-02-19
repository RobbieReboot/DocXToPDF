using System;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Core
{
    public interface IPdfRenderableObject
    {
        string Render();
        byte[] RenderBytes(long filePos, out int size);
    }
}
