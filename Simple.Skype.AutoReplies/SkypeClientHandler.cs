using SKYPE4COMLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Simple.Skype.AutoReplies
{
    public class SkypeClientHandler : ISkypeClientHandler
    {
        private readonly string _clientHandler;

        private readonly SkypeClass _skype;

        private TUserStatus _currentUserStatus;

        public SkypeClientHandler(string clientHandler, SkypeClass skypeClient)
        {
            if (clientHandler == null)
            {
                throw new ArgumentNullException(nameof(clientHandler));
            }

            if (skypeClient == null)
            {
                throw new ArgumentNullException(nameof(skypeClient));
            }

            _clientHandler = clientHandler;
            _skype = skypeClient;
        }

        public void Init()
        {
            _skype.UserStatus += UserStatusChanged;
            _skype.MessageStatus += MessageReceived;
            _currentUserStatus = _skype.CurrentUserStatus;
        }

        private void MessageReceived(ChatMessage pMessage, TChatMessageStatus Status)
        {
            if (_currentUserStatus != TUserStatus.cusOnline && Status == TChatMessageStatus.cmsReceived)
            {
                _skype.SendMessage(pMessage.Sender.Handle, Properties.Settings.Default.AutoReplyMessage);
            }
        }

        private void UserStatusChanged(TUserStatus status)
        {
            _currentUserStatus = status;
        }

        public void Dispose()
        {
            if (_skype != null)
            {
                Marshal.ReleaseComObject(_skype);
            }
        }
    }
}
