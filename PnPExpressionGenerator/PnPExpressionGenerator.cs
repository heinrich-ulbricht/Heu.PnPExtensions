using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.IO;
using System.Globalization;

namespace PnPExtensions
{
    public class PnPExpressionGenerator
    {
        ScriptOptions options = null;
        ScriptOptions Options
        {
            get
            {
                if (options == null)
                { 
                    options = InitOptions();
                }
                return options;
            }
        }

        private ScriptOptions InitOptions()
        {
            // let's see if those assemblies are enough to reference...
            var partialAssemblyNames = new string[] { "Microsoft.SharePoint" };
            var referencedAssemblies = new List<Assembly>();
            foreach (var namePart in partialAssemblyNames)
            {
                var list = AppDomain.CurrentDomain.GetAssemblies().Where(a => a.FullName.StartsWith(namePart));
                referencedAssemblies.AddRange(list);
            }
            referencedAssemblies.Add(typeof(Expression).Assembly);

            var options = ScriptOptions.Default.AddReferences(referencedAssemblies);
            options.AddImports("Microsoft.SharePoint.Client", "System.Linq", "System.Linq.Expressions");
            return options;
        }

        private PropertyInfo GetPropertyByName(Type type, string propName)
        {
            var objectProps = type.GetProperties();
            return objectProps.Where(ra => ra.Name.Equals(propName, StringComparison.InvariantCultureIgnoreCase)).SingleOrDefault();
        }

        private Type GetGenericTypeFromIQueryableInterface(Type type)
        {
            // get interfaces os member, e.g. RoleAssignments - we'll search for IQueryable<T> and want to know what T is, in this case "RoleAssignment"
            List<Type> genTypes = new List<Type>();
            foreach (Type intType in type.GetInterfaces())
            {
                if (intType.IsGenericType && intType.GetGenericTypeDefinition() == typeof(IQueryable<>))
                {
                    genTypes.Add(intType.GetGenericArguments()[0]);
                }
            }

            return genTypes.FirstOrDefault();
        }

        public string GenerateExpressionCode(Type T, string membersWithDot)
        {
            var currentVariable = 'a';
            var members = membersWithDot.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            var typeFromCollection = GetGenericTypeFromIQueryableInterface(T);
            string finalCode;
            if (typeFromCollection == null)
            {
                finalCode = $"{currentVariable} => ___";
            }
            else
            {
                // if calling for collection like on a ListItemCollection
                finalCode = $"{currentVariable} => Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include({currentVariable}, {++currentVariable} => ___)";
            }

            var currentType = typeFromCollection ?? T;

            var currentCode = $"{currentVariable}";
            foreach (var m in members)
            {
                var prop = GetPropertyByName(currentType, m);
                if (prop == null)
                {
                    // if property is not found assume that it is retrievable via clientObject["memberName"], e.g. for ListItem["FileRef"]
                    currentCode += $"[\"{m}\"]";
                    break;
                }
                currentType = prop.PropertyType;

                currentCode += $".{m}";
                typeFromCollection = GetGenericTypeFromIQueryableInterface(currentType);
                // check if we got collection
                if (typeFromCollection != null)
                {
                    currentType = typeFromCollection;

                    currentVariable++;
                    currentCode = $"Microsoft.SharePoint.Client.ClientObjectQueryableExtension.Include({currentCode}, {currentVariable} => ___)";

                    finalCode = finalCode.Replace("___", currentCode);
                    currentCode = $"{currentVariable}";
                }
            }

            finalCode = finalCode.Replace("___", $"{currentCode}");
            return finalCode;
        }

        public Expression<Func<T, object>> GetExpression<T>(T clientObject, string memberNameWithDot)
        {
            var code = GenerateExpressionCode(typeof(T), memberNameWithDot);
            return GenerateExpressionWhileUsingAlreadyLoadedTypes(clientObject, code);
        }

        public Expression<Func<T, object>>[] GetExpressions<T>(T clientObject, params string[] memberNamesWithDot)
        {
            var expressions = new List<Expression<Func<T, object>>>();
            foreach (var m in memberNamesWithDot)
            {
                // todo: make the expressions being generated parallelly again
                expressions.Add(GetExpression(clientObject, m));
            }
            return expressions.ToArray();        
        }

        public Expression<Func<T, object>> GetWhereExpression<T>(T clientObject, string filterExpression)
        {
            var code = $"reallyuniquevariablename_heu => System.Linq.Queryable.Where(reallyuniquevariablename_heu, {filterExpression})";
            return GenerateExpressionWhileUsingAlreadyLoadedTypes(clientObject, code);
        }

        private Expression<Func<T, object>> GenerateExpressionWhileUsingAlreadyLoadedTypes<T>(T clientObject, string expressionCode)
        {
            var incomingType = typeof(T);
            // todo: check that clientObject is decendant of ClientObject

            // this disables type constraint "T : ClientObject" so we are completely independent from a specific version of the client libraries
            var script = $"System.Linq.Expressions.Expression<System.Func<{incomingType.FullName}, object>> GetExpression() {{ return {expressionCode}; }} return GetExpression();";
            var result = CSharpScript
                .Create(script, Options)
                .RunAsync().GetAwaiter().GetResult();

            if (result.Exception != null)
            {
                throw new Exception($"Error while generating expression for code snippet '{expressionCode}'", result.Exception);
            }

            return (Expression < Func<T, object> > )result.ReturnValue;
        }
    }
}