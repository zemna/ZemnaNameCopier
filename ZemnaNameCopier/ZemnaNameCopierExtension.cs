using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ZemnaNameCopier.Properties;

namespace ZemnaNameCopier
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Class, "*")]
    [COMServerAssociation(AssociationType.Directory)]
    [COMServerAssociation(AssociationType.Class, "Directory\\Background")]
    [COMServerAssociation(AssociationType.Drive)]
    public class ZemnaNameCopierExtension : SharpContextMenu
    {   
        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();

            var itemCopyAbsolutePath = new ToolStripMenuItem
            {
                Text = "Copy &Absolute Path(s)",
                Image = Resources.ZemnaNameCopier
            };

            var itemCopyName = new ToolStripMenuItem
            {
                Text = "Copy &Name",
                Image = Resources.ZemnaNameCopier
            };

            itemCopyAbsolutePath.Click += (sender, args) => CopyAbsolutePath();
            itemCopyName.Click += (sender, args) => CopyName();

            menu.Items.Add(itemCopyAbsolutePath);
            menu.Items.Add(itemCopyName);

            return menu;
        }

        private void CopyAbsolutePath()
        {
            var paths = new StringBuilder();

            var selectedPaths = new List<String>(SelectedItemPaths);
            if (!string.IsNullOrEmpty(FolderPath))
                selectedPaths.Add(FolderPath);

            foreach (var itemPath in selectedPaths)
            {
                paths.AppendLine(itemPath);
            }

            Clipboard.SetText(paths.ToString());
        }

        private void CopyName()
        {
            var paths = new StringBuilder();

            var selectedPaths = new List<String>(SelectedItemPaths);
            if (!string.IsNullOrEmpty(FolderPath))
                selectedPaths.Add(FolderPath);

            foreach (var itemPath in selectedPaths)
            {
                paths.AppendLine(Path.GetFileName(itemPath));
            }

            Clipboard.SetText(paths.ToString());
        }
    }
}
