using Simple.Skype.AutoReplies.Properties;
using SKYPE4COMLib;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Simple.Skype.AutoReplies
{
    public class SkypeAutoReplyApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;

        private ConcurrentDictionary<string, ISkypeClientHandler> _skypeClients;

        private bool _continue;

        private CancellationTokenSource _processCancellationToken;

        private readonly object _lockObject = new object();

        public SkypeAutoReplyApplicationContext()
        {
            _skypeClients = new ConcurrentDictionary<string, ISkypeClientHandler>();

            trayIcon = new NotifyIcon()
            {
                Icon = Resources.SkypeAutoReplyIcon,
                ContextMenu = new ContextMenu(new MenuItem[] {
                new MenuItem("Exit", Exit)
            }),
                Visible = true
            };

            Init();
        }

        void Exit(object sender, EventArgs e)
        {
            trayIcon.Visible = false;

            Flush();

            System.Windows.Forms.Application.Exit();
        }


        protected void Init()
        {
            try
            {
                lock (_lockObject)
                {
                    _processCancellationToken = new CancellationTokenSource();

                    _continue = true;

                    Task.Factory.StartNew(
                        () =>
                        {
                            _processCancellationToken.Token.ThrowIfCancellationRequested();
                            Process();
                        },
                        _processCancellationToken.Token,
                         TaskCreationOptions.DenyChildAttach,
                         TaskScheduler.Current);
                }
            }
            catch (Exception exe) when (exe is TaskCanceledException || exe is ThreadAbortException)
            {

            }
            catch (Exception exe)
            {
                throw;
            }
        }

        private void Flush()
        {
            _processCancellationToken.Cancel();

            _continue = false;

            lock (_lockObject)
            {
                foreach (var _skypeClient in _skypeClients)
                {
                    _skypeClient.Value.Dispose();
                }
            }
        }

        private void Process()
        {
            while (_continue)
            {
                SkypeClass _skype = new SkypeClass();

                _skype.Attach(7, false);

                if (_skype.AttachmentStatus == TAttachmentStatus.apiAttachSuccess && !_skypeClients.ContainsKey(_skype.CurrentUserHandle))
                {
                    var clientHandler = new SkypeClientHandler(_skype.CurrentUserHandle, _skype);

                    if (_skypeClients.TryAdd(_skype.CurrentUserHandle, clientHandler))
                    {
                        Task.Factory.StartNew(() => clientHandler.Init());
                    }
                }
                else
                {
                    Marshal.ReleaseComObject(_skype);
                }

                Thread.Sleep(Properties.Settings.Default.SleepTimeForClientAdd);
            }
        }
        protected override void Dispose(bool disposing)
        {
            Flush();
            base.Dispose(disposing);
        }
    }
}
