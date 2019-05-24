﻿using Panacea.Core;
using Panacea.Modularity.Hardware;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Panacea.Modules.Remedi
{
    public class RemediPlugin : IHardwarePlugin
    {
        // only one instance even if multiple plugin instances
        static Remedi _manager;
        static readonly object _lock = new object();
        private readonly PanaceaServices _core;

        public RemediPlugin(PanaceaServices core)
        {
            _core = core;
        }

        public Task BeginInit()
        {
            return Task.CompletedTask;
        }

        public void Dispose()
        {
            
        }

        public Task EndInit()
        {
            return Task.CompletedTask;
        }

        public IHardwareManager GetHardwareManager()
        {
            if(_manager == null)
            {
                lock (_lock)
                {
                    if(_manager == null)
                    {
                        _manager = new Remedi(_core.Logger);
                        _manager.Start();
                    }
                }
            }
            return _manager;
        }

        public Task Shutdown()
        {
            _manager?.Stop();
            return Task.CompletedTask;
        }
    }
}
