using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Vivit_Control_Center.Views;

namespace Vivit_Control_Center.Views.Modules
{
    public class BaseSimpleModule : UserControl, IModule
    {
        private bool _signaled;
        private readonly TaskCompletionSource<bool> _tcs =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        public event EventHandler LoadCompleted;
        public Task LoadCompletedTask => _tcs.Task;
        public FrameworkElement View => this;

        public BaseSimpleModule()
        {
            Loaded += (_, __) => SignalLoadedOnce();
        }

        public virtual Task PreloadAsync()
        {
            SignalLoadedOnce();
            return _tcs.Task;
        }

        public virtual void SetVisible(bool visible)
        {
            Visibility = visible ? Visibility.Visible : Visibility.Hidden;
            IsHitTestVisible = visible;
        }

        protected void SignalLoadedOnce()
        {
            if (_signaled) return;
            _signaled = true;
            _tcs.TrySetResult(true);
            try { LoadCompleted?.Invoke(this, EventArgs.Empty); } catch { }
        }
    }
}



