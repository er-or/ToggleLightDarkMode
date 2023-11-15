using System.Collections;

namespace ToggleLightDarkMode {
    partial class Installer1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        public override void Commit(IDictionary savedState) {
            base.Commit(savedState);
            System.Diagnostics.Process.Start(System.IO.Path.GetDirectoryName(this.Context.Parameters["AssemblyPath"]) + "\\Toggle Light Dark Mode.exe");
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            components = new System.ComponentModel.Container();
        }

        #endregion
    }
}