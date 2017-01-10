using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a set of methods for interacting with an Excel <see cref="Excel.ListObject"/>.
    /// </summary>
    /// <remarks>
    /// By moving as much <see cref="Excel.ListObject"/> interaction as we can into classes 
    /// which implement this interface we are able to maximize the amount of unit testing within
    /// the project.
    /// </remarks>
    public interface IListObjHandling
    {
        /// <summary>
        /// Gets the underlying Excel table of interest.
        /// </summary>
        Excel.ListObject XLTable { get; }

        /// <summary>
        /// Gets a value indicating whether or not the underlying <see cref="Excel.ListObject"/> exists.
        /// </summary>
        /// <remarks>This property is created so that we can catch cases where the user 
        /// has managed to delete or to rename the Excel table.</remarks>
        bool ListObjectExists { get; }

        /// <summary>
        /// Gets a value indicating whether or not the underlying <see cref="Excel.ListObject"/> 
        /// has yet been populated with any data.
        /// </summary>
        bool ListObjectHasData { get; }

        /// <summary>
        /// Gets the <see cref="iClickerQuizPts.WshListobjPair"/> for the class, 
        /// indicating the name of the <see cref="Excel.ListObject"/> and the name of its
        /// parent <see cref="Excel.Worksheet"/>.
        /// </summary>
        WshListobjPair WshListObjPair { get;  }
    }
}
