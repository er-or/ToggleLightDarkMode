﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ToggleLightDarkMode.languages {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class lang_fr {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal lang_fr() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("ToggleLightDarkMode.languages.lang_fr", typeof(lang_fr).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Quitter.
        /// </summary>
        internal static string exit {
            get {
                return ResourceManager.GetString("exit", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Ce programme est déjà en cours d&apos;exécution.
        /// </summary>
        internal static string program_already_running {
            get {
                return ResourceManager.GetString("program_already_running", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Passer en mode sombre.
        /// </summary>
        internal static string switch_to_dark_mode {
            get {
                return ResourceManager.GetString("switch_to_dark_mode", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Passer en mode clair.
        /// </summary>
        internal static string switch_to_light_mode {
            get {
                return ResourceManager.GetString("switch_to_light_mode", resourceCulture);
            }
        }
    }
}
