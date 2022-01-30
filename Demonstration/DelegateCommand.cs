using System;
using System.Windows.Input;

namespace Demonstration
{
    public class DelegateCommand : ICommand
    {
        private readonly System.Action? executeMethod = null;
        private readonly Func<bool>? canExecuteMethod = null;

        public event EventHandler? CanExecuteChanged;

        public DelegateCommand(System.Action executeMethod)
            : this(executeMethod, null)
        {
        }

        public DelegateCommand(System.Action? executeMethod, Func<bool>? canExecuteMethod)
        {
            if (executeMethod == null)
            {
                throw new ArgumentNullException(nameof(executeMethod));
            }

            this.executeMethod = executeMethod;
            this.canExecuteMethod = canExecuteMethod;
        }

        public bool CanExecute()
        {
            if (canExecuteMethod != null)
            {
                return canExecuteMethod();
            }

            return true;
        }

        public void Execute()
        {
            executeMethod?.Invoke();
        }

        bool ICommand.CanExecute(object? parameter)
        {
            return CanExecute();
        }

        void ICommand.Execute(object? parameter)
        {
            Execute();
        }
    }
}
