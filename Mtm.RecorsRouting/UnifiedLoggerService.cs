// --------------------------------------------------------------------------------------------------------------------
// <copyright file="UnifiedLoggerService.cs" company="Montrium">
//   MIT License
// </copyright>
// <summary>
//   The unified logger service.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting
{
    using System;
    using System.Collections.Generic;

    using Microsoft.SharePoint.Administration;

    /// <summary>
    /// The unified logger service.
    /// </summary>
    [CLSCompliant(false)]
    public class UnifiedLoggerService : SPDiagnosticsServiceBase
    {
        #region AttributesOrProperties
        /// <summary>The service name.</summary>
        private static string ServiceName = "Montrium Logging Service";

        /// <summary>The diagnostics area name.</summary>
        private static string DiagnosticsAreaName = "Montrium Solutions";

        /// <summary>The category.</summary>
        private static string Category = "Montrium RUBi_Methods";

        /// <summary>The event id.</summary>
        private static int EventId = 9191;

        /// <summary>Gets The current.</summary>
        private static UnifiedLoggerService current;

        /// <summary>Gets the current.</summary>
        public static UnifiedLoggerService Current
        {
            get
            {
                return current ?? (current = new UnifiedLoggerService());
            }
        }
        #endregion

        #region Constructors
        /// <summary>Prevents a default instance of the <see cref="UnifiedLoggerService"/> class from being created.</summary>
        private UnifiedLoggerService()
            : base(ServiceName, SPFarm.Local)
        {
        }
        #endregion

        #region Destructor
        #endregion

        #region PublicMethods
        /// <summary>Writes a High level message to the SharePoint ULS only.</summary>
        /// <param name="message">The message.</param>
        public static void High(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;
            WriteLog(TraceSeverity.High, message);
        }

        /// <summary>Writes a Medium level message to the SharePoint ULS only.</summary>
        /// <param name="message">The message.</param>
        public static void Medium(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;
            WriteLog(TraceSeverity.Medium, message);
        }

        /// <summary>Writes a Low level message to the SharePoint ULS only (i.e. Most Verbose, or most detailed).</summary>
        /// <param name="message">The message.</param>
        public static void Low(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;
            WriteLog(TraceSeverity.Verbose, message);
        }

        /// <summary>
        /// Unexpecteds the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        public static void Unexpected(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;
            WriteLog(TraceSeverity.Unexpected, message);
        }
        #endregion

        #region OverrideMethods
        /// <summary>
        /// The provide areas.
        /// </summary>
        /// <returns>
        /// The System.Collections.Generic.IEnumerable`1[T -&gt; Microsoft.SharePoint.Administration.SPDiagnosticsArea].
        /// </returns>
        protected override IEnumerable<SPDiagnosticsArea> ProvideAreas()
        {
            // provide the category with default severities 
            List<SPDiagnosticsCategory> categories = new List<SPDiagnosticsCategory>
                {
                    new SPDiagnosticsCategory(Category, TraceSeverity.Medium, EventSeverity.Information)
                };

            yield return new SPDiagnosticsArea(DiagnosticsAreaName, 0, 0, false, categories);
        }

        #endregion

        #region PrivateMethods
        /// <summary>
        /// Writes the log.
        /// </summary>
        /// <param name="traceSeverity">The trace severity.</param>
        /// <param name="message">The message.</param>
        private static void WriteLog(TraceSeverity traceSeverity, string message)
        {
            if (traceSeverity != TraceSeverity.None)
            {
                try
                {
                    SPDiagnosticsCategory spDiagnosticsCategory = UnifiedLoggerService.Current.Areas[DiagnosticsAreaName].Categories[Category];
                    UnifiedLoggerService.Current.WriteTrace((uint)EventId, spDiagnosticsCategory, traceSeverity, message);
                }
                catch
                {
                }
            }
        }

        #endregion
    }
}