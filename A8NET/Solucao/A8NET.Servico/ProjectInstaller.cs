using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;

namespace A8NET.Servico
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }
    }
}