using AcGraphicToFrame.Exceptions;

namespace AcGraphicToFrame.Helpers
{
    internal static class FormatHelper
    {
        internal static string GetFormatValue(double modelHeight, double modelWidth, double scale)
        {
            if (((Constants.HeightA3 - 20) * scale) > modelHeight && (Constants.WidthA3 - 20) * scale > modelWidth)
            {
                return "A3";
            }
            else if (((Constants.HeightA2 - 20) * scale) > modelHeight && (Constants.WidthA2 - 20) * scale > modelWidth)
            {
                return "A2";
            }
            else if (((Constants.HeightA1 - 20) * scale) > modelHeight && (Constants.WidthA1 - 20) * scale > modelWidth)
            {
                return "A1";
            }
            else
            {
                throw new FormatNotFoundException($"Unable to find format name for this drawing: HEIGHT - {modelHeight}, WIDTH - {modelWidth}, SCALE - {scale}");
            };
        }
    }
}
