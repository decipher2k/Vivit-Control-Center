using System;
using System.Threading.Tasks;
using System.Windows;

namespace Vivit_Control_Center.Views
{
    public interface IModule
    {
        event EventHandler LoadCompleted;
        Task LoadCompletedTask { get; }
        Task PreloadAsync();
        void SetVisible(bool visible);
        FrameworkElement View { get; }
    }
}