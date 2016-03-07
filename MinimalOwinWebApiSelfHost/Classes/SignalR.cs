using Microsoft.AspNet.SignalR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinimalOwinWebApiSelfHost.Classes
{
    public class MyHub : Hub
    {
        public void Send(string message)
        {
            Clients.All.addMessage(message);
        }
    }
}
