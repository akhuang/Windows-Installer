using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace CustomBootstrapper.Core
{
    public class CommandHandler : ICommand
    {
        private Action _action;
        private Func<bool> _canExecute;

        public CommandHandler(Action action)
            : this(action, null)
        {
        }

        public CommandHandler(Action action, Func<bool> canExecute)
        {
            _action = action;
            _canExecute = canExecute;
        }



        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute();
        }

        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            if (CanExecute(parameter))
            {
                _action();
            }
        }
    }
}
