using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts.Itfs
{
    /// <summary>
    /// Provides the interface for classes providing wrapper access to
    /// <see cref="Excel.Names"/>.
    /// </summary>
    public interface INamedRangeWrapping
    {
        /// <summary>
        /// Tests whether a specified workbook-scoped <see cref="Excel.Name"/> both exists 
        /// and refers to a valid <see cref="Excel.Range"/>.
        /// </summary>
        /// <param name="rngNm">The name of the <see cref="Excel.Range"/>.</param>
        /// <returns><c>true</c> if the <see cref="Excel.Name"/> both exists and refers to a valid 
        /// <see cref="Excel.Range"/>; otherwise <c>false</c>.</returns>
        bool WorkbookScopedRangeExists(string rngNm);

        /// <summary>
        /// Tests whether a specified worksheet-scoped <see cref="Excel.Name"/> both exists 
        /// and refers to a valid <see cref="Excel.Range"/>.
        /// </summary>
        /// <param name="wshNm">The name of the parent <see cref="Excel.Worksheet"/>.</param>
        /// <param name="rngNm">The name of the <see cref="Excel.Range"/>.</param>
        /// <returns><c>true</c> if the <see cref="Excel.Name"/> both exists and refers to a valid 
        /// <see cref="Excel.Range"/>; otherwise <c>false</c>.</returns>
        bool WorksheetScopedRangeExists(string wshNm, string rngNm);
    }
}