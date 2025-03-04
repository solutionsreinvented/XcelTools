using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;

namespace XcelTools.XpLore.Services
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class SectionsProvider
    {
        private static SectionsProvider _instance;

        private SectionsProvider()
        {

        }

        public static SectionsProvider Instance()
        {
            return _instance ?? (_instance = new SectionsProvider());
        }

        public void ShowGreeting()
        {
            _ = MessageBox.Show("I was asked to greet you by XcelTools. Hello! Good morning");
        }

        public List<int> GetIntCollection()
        {
            return new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
        }

    }
}
