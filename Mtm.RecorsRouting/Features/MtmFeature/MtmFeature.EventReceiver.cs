// --------------------------------------------------------------------------------------------------------------------
// <copyright file="MtmFeature.EventReceiver.cs" company="Montrium">
//   MIT Licence
// </copyright>
// <summary>
//   This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting.Features.MtmFeature
{
    using System.Runtime.InteropServices;
    using Microsoft.SharePoint;


    /// <summary>This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.</summary>
    /// <remarks>The GUID attached to this class may be used during packaging and should not be modified.</remarks>
    [Guid("e82d7f63-c96d-4fbd-b7a4-3a7ff8a31b27")]
    public class Feature1EventReceiver : SPFeatureReceiver
    {
        /// <summary>Feature Activated.</summary>
        /// <param name="properties">The properties.</param>
        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            const string Url = "http://dev2010/sites/rc/";

            RecordCentreManager rcm = new RecordCentreManager(Url);
        }

        // Uncomment the method below to handle the event raised before a feature is deactivated.

        /// <summary>Feature Deactivating.</summary>
        /// <param name="properties">The properties.</param>
        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            using (SPWeb spWeb = properties.Feature.Parent as SPWeb)
            {
                if (spWeb != null)
                {
                    SPContentType myContentType = spWeb.ContentTypes["New Announcements"];
                    spWeb.ContentTypes.Delete(myContentType.Id);
                    spWeb.Fields["Team Project"].Delete();
                }
            }
        }

        // Uncomment the method below to handle the event raised after a feature has been installed.

        // public override void FeatureInstalled(SPFeatureReceiverProperties properties)
        // {
        // }

        // Uncomment the method below to handle the event raised before a feature is uninstalled.

        // public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        // {
        // }

        // Uncomment the method below to handle the event raised when a feature is upgrading.

        // public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
        // {
        // }
    }
}
