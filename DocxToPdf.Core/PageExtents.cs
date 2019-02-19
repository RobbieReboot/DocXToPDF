namespace DocxToPdf.Core
{
    public struct PageExtents
    {
        public uint xWidth;
        public uint yHeight;
        public uint leftMargin;
        public uint rightMargin;
        public uint topMargin;
        public uint bottomMargin;

        public PageExtents(uint width, uint height, uint left = 0, uint top = 0, uint right = 0, uint bottom = 0)
        {
            xWidth = width;
            yHeight = height;
            leftMargin = left;
            rightMargin = right;
            topMargin = top;
            bottomMargin = bottom;
        }

        public void SetMargins(uint left, uint top, uint right, uint bottom)
        {
            leftMargin = left;
            rightMargin = right;
            topMargin = top;
            bottomMargin = bottom;
        }
    }
}