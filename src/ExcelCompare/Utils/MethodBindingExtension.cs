using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;

namespace ExcelCompare.Utils
{
    public class MethodBindingExtension : MarkupExtension
    {
        private readonly object[] args;

        public MethodBindingExtension(object arg) : this(new object[] { arg }) { }
        public MethodBindingExtension(object arg0, object arg1) : this(new object[] { arg0, arg1 }) { }
        public MethodBindingExtension(object arg0, object arg1, object arg2) : this(new object[] { arg0, arg1, arg2 }) { }
        public MethodBindingExtension(object arg0, object arg1, object arg2, object arg3) : this(new object[] { arg0, arg1, arg2, arg3 }) { }
        public MethodBindingExtension(object arg0, object arg1, object arg2, object arg3, object arg4) : this(new object[] { arg0, arg1, arg2, arg3, arg4 }) { }

        private MethodBindingExtension(object[] args)
        {
            this.args = args;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            if (serviceProvider.GetService(typeof(IProvideValueTarget)) is IProvideValueTarget provideValueTarget &&
                provideValueTarget.TargetObject is DependencyObject dependencyObject &&
                provideValueTarget.TargetProperty is EventInfo eventInfo)
            {
                if (args.Length <= 0)
                {
                    return this;
                }
                var markupExpression = MarkupExpression.Load(dependencyObject, args);
                if (markupExpression == null)
                {
                    return this;
                }
                return CreateEventHandler(markupExpression, eventInfo.EventHandlerType);
            }
            return this;
        }

        private Delegate CreateEventHandler(MarkupExpression markupExpression, Type eventHandlerType)
        {
            EventHandler handler = (sender, e) =>
            {
                var viewModelType = markupExpression.TargetObject?.GetType();
                var methodInfo = viewModelType?.GetMethod(markupExpression.MethodName);
                if (methodInfo == null) { return; }
                var instance = System.Linq.Expressions.Expression.Constant(markupExpression.TargetObject);
                var expressionArgs = markupExpression.Args
                .Select(arg => arg is EventArgsExtension ? e : arg)
                .Select(arg => arg is OpenFileDialogExtension openFileDialog ? openFileDialog.SelectFilePath : null)
                .Select(arg => System.Linq.Expressions.Expression.Constant(arg));

                try
                {
                    var callExpression = System.Linq.Expressions.Expression.Call(
                            instance,
                            methodInfo,
                            expressionArgs);
                    var compile = System.Linq.Expressions.Expression.Lambda<Action>(callExpression).Compile();
                    compile.Invoke();
                }
                catch { }
            };
            return Delegate.CreateDelegate(eventHandlerType, handler.Target, handler.Method);
        }
    }

    public class EventArgsExtension : MarkupExtension
    {
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }

    public class OpenFileDialogExtension : MarkupExtension
    {
        private readonly string filter;

        public string SelectFilePath
        {
            get
            {
                OpenFileDialog openFileDialog = new() { Filter = filter };
                if (openFileDialog.ShowDialog() == true)
                {
                    return openFileDialog.FileName;
                }
                return null;
            }
        }

        public OpenFileDialogExtension(string filter)
        {
            this.filter = filter;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }

    internal abstract class MarkupExpression
    {
        private static readonly List<DependencyProperty> storageProperties = new List<DependencyProperty>();
        public static MarkupExpression Load(DependencyObject dependencyObject, object[] args)
        {
            if (args.Length <= 0)
            {
                return null;
            }
            {
                if (args.Length >= 1 &&
                  args[0] is string methodName &&
                 !string.IsNullOrEmpty(methodName) &&
                  dependencyObject is FrameworkElement frameworkElement
                  )
                {
                    return new MarkupExpressionDataContext(frameworkElement, methodName, args.Skip(1));
                }
            }
            {
                if (args.Length >= 2 &&
                 args[0] is MarkupExtension markupExtension &&
                 AsDependencyProperty(dependencyObject, markupExtension) is DependencyProperty dependencyProperty &&
                 args[1] is string methodName &&
                !string.IsNullOrEmpty(methodName))
                {
                    return new MarkupExpressionFrameworkElement(dependencyObject, dependencyProperty, methodName, args.Skip(2));
                }
            }
            return null;
        }

        private static DependencyProperty AsDependencyProperty(DependencyObject dependencyObject, object arg)
        {
            var dependencyProperty = storageProperties.FirstOrDefault(p => dependencyObject.ReadLocalValue(p) == DependencyProperty.UnsetValue);
            if (dependencyProperty == null)
            {
                dependencyProperty = DependencyProperty.RegisterAttached(dependencyObject.GetType().Name + storageProperties.Count,
                    typeof(object),
                    dependencyObject.GetType(),
                    new PropertyMetadata());
                storageProperties.Add(dependencyProperty);
            }
            if (arg is MarkupExtension markupExtension)
            {
                var resolvedValue = markupExtension.ProvideValue(new ServiceProvider(dependencyObject, dependencyProperty));
                dependencyObject.SetValue(dependencyProperty, resolvedValue);
            }
            else
            {
                dependencyObject.SetValue(dependencyProperty, arg);
            }
            return dependencyProperty;
        }

        public abstract object TargetObject { get; }
        public abstract string MethodName { get; }
        public abstract IEnumerable<object> Args { get; }

        private MarkupExpression() { }

        private class ServiceProvider : IServiceProvider, IProvideValueTarget
        {
            public object TargetObject { get; }
            public object TargetProperty { get; }

            public ServiceProvider(object targetObject, object targetProperty)
            {
                TargetObject = targetObject;
                TargetProperty = targetProperty;
            }

            public object GetService(Type serviceType)
            {
                return serviceType.IsInstanceOfType(this) ? this : null;
            }
        }

        private class MarkupExpressionDataContext : MarkupExpression
        {
            private readonly FrameworkElement frameworkElement;
            private readonly string methodName;
            private readonly IEnumerable<DependencyProperty> dependencyProperties;

            public override object TargetObject => frameworkElement.DataContext;

            public override string MethodName => methodName;

            public override IEnumerable<object> Args => dependencyProperties.Select(d => frameworkElement.GetValue(d));

            public MarkupExpressionDataContext(FrameworkElement frameworkElement, string methodName, IEnumerable<object> args)
            {
                this.frameworkElement = frameworkElement;
                this.methodName = methodName;
                dependencyProperties = args?.Select(arg => AsDependencyProperty(frameworkElement, arg))?.ToArray();
            }
        }

        private class MarkupExpressionFrameworkElement : MarkupExpression
        {
            private readonly DependencyObject dependencyObject;
            private readonly DependencyProperty dependencyProperty;
            private readonly string methodName;
            private readonly IEnumerable<DependencyProperty> dependencyProperties;

            public override object TargetObject => dependencyObject.GetValue(dependencyProperty);

            public override string MethodName => methodName;

            public override IEnumerable<object> Args => dependencyProperties.Select(d => dependencyObject.GetValue(d));

            public MarkupExpressionFrameworkElement(DependencyObject dependencyObject, DependencyProperty dependencyProperty, string methodName, IEnumerable<object> args)
            {
                this.dependencyObject = dependencyObject;
                this.dependencyProperty = dependencyProperty;
                this.methodName = methodName;
                dependencyProperties = args?.Select(arg => AsDependencyProperty(dependencyObject, arg))?.ToArray();
            }
        }
    }
}
