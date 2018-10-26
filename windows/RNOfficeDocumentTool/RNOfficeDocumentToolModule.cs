using ReactNative.Bridge;
using System;
using System.Collections.Generic;
using Windows.ApplicationModel.Core;
using Windows.UI.Core;

namespace Office.Document.Tool.RNOfficeDocumentTool
{
    /// <summary>
    /// A module that allows JS to share data.
    /// </summary>
    class RNOfficeDocumentToolModule : NativeModuleBase
    {
        /// <summary>
        /// Instantiates the <see cref="RNOfficeDocumentToolModule"/>.
        /// </summary>
        internal RNOfficeDocumentToolModule()
        {

        }

        /// <summary>
        /// The name of the native module.
        /// </summary>
        public override string Name
        {
            get
            {
                return "RNOfficeDocumentTool";
            }
        }
    }
}
