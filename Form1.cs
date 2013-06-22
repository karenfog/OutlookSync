using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSync
{
	public partial class Form1 : Form
	{
		public Form1() {
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e) {
			var root = GetRootFolder();
			var node = new TreeNode(root.Name);

			treeView1.Nodes.Add(node);
			RenderTreeView(root, node);
		}

		private Store GetStoreByPath(NameSpace ns, string path) {
			foreach (Outlook.Store store in ns.Stores) {
				if (store.FilePath.Equals(path, StringComparison.InvariantCulture))
					return store;
			}

			return null;
		}

		private Outlook.Folder GetRootFolder() {
			var app = new Outlook.Application();
			var @namespace = app.Application.GetNamespace("MAPI");
			var path = @"D:\Karen\OutlookSync\backup.pst";

			@namespace.AddStore(path);

			return GetStoreByPath(@namespace, path).GetRootFolder() as Outlook.Folder;
		}

		private void RenderTreeView(Folder root, TreeNode node) {
			foreach (Outlook.Folder folder in root.Folders) {
				TreeNode childNode = new TreeNode(folder.Name);

				node.Nodes.Add(childNode);
				RenderTreeView(folder, childNode);
			}
		}
	}
}
