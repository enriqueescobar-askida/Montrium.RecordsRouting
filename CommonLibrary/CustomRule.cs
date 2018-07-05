// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CustomRule.cs" company="Montrium">
//   MIT Licence
// </copyright>
// <summary>
//   Defines the CustomRecordRule type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting.CommonLibrary
{
    using System;
    using System.IO;

    /// <summary>Custom RecordRule.</summary>
    public class CustomRule : IDisposable
    {
        #region fields
        /// <summary>is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>Initializes a new instance of the <see cref="CustomRule"/> class.</summary>
        /// <param name="ruleName">The rule name.</param>
        /// <param name="ruleDesc">The rule desc.</param>
        /// <param name="ruleLibName">The rule lib name.</param>
        /// <param name="ruleFolder">The rule folder.</param>
        /// <param name="ruleContentTypeName">The rule content type name.</param>
        public CustomRule(
            string ruleName, string ruleDesc, string ruleLibName, string ruleFolder, string ruleContentTypeName)
        {
            this.RuleName = ruleName;
            this.RuleDescription = ruleDesc;
            this.RuleLibraryName = ruleLibName;
            this.RuleFolder = Path.Combine(ruleLibName, ruleFolder);
            this.RuleContentTypeName = ruleContentTypeName;
        }
        #endregion

        #region DestructorDisposable
        /// <summary>Finalizes an instance of the <see cref="CustomRule"/> class.</summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class
        ~CustomRule()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region Attributes
        /// <summary>Gets the name of the rule.</summary>
        public string RuleName { get; internal set; }

        /// <summary>Gets the rule description.</summary>
        public string RuleDescription { get; internal set; }

        /// <summary>Gets the name of the rule library.</summary>
        public string RuleLibraryName { get; internal set; }

        /// <summary>Gets the rule folder.</summary>
        public string RuleFolder { get; internal set; }

        /// <summary>Gets the name of the rule content type.</summary>
        public string RuleContentTypeName { get; internal set; }
        #endregion

        #region PublicOverride
        /// <summary>
        /// Returns a <see cref="System.String"/> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String"/> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            string s = String.Empty;
            return s + "\n";
        }
        #endregion

        #region PublicDisposable
        /// <summary>Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.</summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region PrivateDisposable
        /// <summary>Dispose Finally.</summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// true to release both managed and unmanaged resources; false to release only unmanaged resources.
        /// <param name="isDisposing">The is disposing.</param>
        private void Dispose(bool isDisposing)
        {
            // Check if Dispose has been called
            if (!this.isDisposed)
            {
                // dispose managed and unmanaged resources
                if (isDisposing)
                {
                    // managed resources clean
                    this.RuleName = this.RuleDescription = this.RuleLibraryName =
                    this.RuleFolder = this.RuleContentTypeName = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion
    }
}
