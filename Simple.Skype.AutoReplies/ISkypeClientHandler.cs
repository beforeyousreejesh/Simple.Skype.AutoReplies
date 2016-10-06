using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Simple.Skype.AutoReplies
{
    public interface ISkypeClientHandler : IDisposable
    {
        void Init();
    }
}