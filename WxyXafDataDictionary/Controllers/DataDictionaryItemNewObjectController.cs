using DevExpress.ExpressApp;
using DevExpress.ExpressApp.SystemModule;
using DevExpress.Persistent.Base;
using System;
using System.Linq;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Wxy.Xaf.DataDictionary;

namespace Wxy.Xaf.DataDictionary.Controllers;

/// <summary>
/// 处理从下拉列表中新建DataDictionaryItem的控制器
/// 确保新建的DataDictionaryItem关联到正确的DataDictionary
/// </summary>
public class DataDictionaryItemNewObjectController : ViewController
{
    private NewObjectViewController _newObjectController;

    public DataDictionaryItemNewObjectController()
    {
        TargetObjectType = typeof(DataDictionaryItem);
    }

    protected override void OnActivated()
    {
        base.OnActivated();
        
        Debug.WriteLine("=== DataDictionaryItemNewObjectController.OnActivated 执行 ===");
        
        _newObjectController = Frame.GetController<NewObjectViewController>();
        if (_newObjectController != null)
        {
            Debug.WriteLine("  找到NewObjectViewController，注册事件处理程序");
            _newObjectController.ObjectCreated += NewObjectController_ObjectCreated;
        }
    }

    private void NewObjectController_ObjectCreated(object sender, ObjectCreatedEventArgs e)
    {
        if (e.CreatedObject is DataDictionaryItem dataDictionaryItem)
        {
            Debug.WriteLine("=== NewObjectController_ObjectCreated 执行 ===");
            Debug.WriteLine($"  新建DataDictionaryItem: {dataDictionaryItem.Name}");
            
            // 获取当前视图的对象空间
            var objectSpace = View.ObjectSpace;
            
            // 从事件参数中获取正确的ObjectSpace，而不是从View.ObjectSpace获取
            var objectSpaceForNewItem = e.ObjectSpace;
            Debug.WriteLine($"  新对象的ObjectSpace: {objectSpaceForNewItem.GetHashCode()}");
            Debug.WriteLine($"  视图的ObjectSpace: {View.ObjectSpace.GetHashCode()}");
            
            // 尝试从当前视图的CollectionSource中获取过滤条件
            if (View is DevExpress.ExpressApp.ListView listView && listView.CollectionSource != null)
            {
                Debug.WriteLine("  当前视图是ListView");
                
                // 检查CollectionSource是否有过滤条件
                var criteriaDictionary = listView.CollectionSource.Criteria;
                if (criteriaDictionary != null && criteriaDictionary.Count > 0)
                {
                    Debug.WriteLine($"  CollectionSource有{criteriaDictionary.Count}个过滤条件");
                    
                    foreach (var key in criteriaDictionary.Keys)
                    {
                        var criteriaOperator = criteriaDictionary[key];
                        Debug.WriteLine($"    过滤条件[{key}]: {criteriaOperator}");
                        
                        // 尝试从过滤条件中提取DataDictionary名称
                        string criteriaString = criteriaOperator.ToString();
                        Debug.WriteLine($"    过滤条件字符串: {criteriaString}");
                        
                        // 使用正则表达式提取DataDictionary名称，支持多种过滤条件格式
                        string dataDictionaryName = ExtractDataDictionaryName(criteriaString);
                        if (!string.IsNullOrEmpty(dataDictionaryName))
                        {
                            Debug.WriteLine($"    提取到DataDictionary名称: {dataDictionaryName}");
                            
                            // 使用新对象的ObjectSpace查找对应的DataDictionary对象，避免SessionMixingException
                            var dataDictionary = objectSpaceForNewItem.FirstOrDefault<DataDictionary>(d => d.Name == dataDictionaryName);
                            if (dataDictionary != null)
                            {
                                // 设置新建的DataDictionaryItem的DataDictionary属性
                                dataDictionaryItem.DataDictionary = dataDictionary;
                                Debug.WriteLine($"    成功设置DataDictionary: {dataDictionary.Name}");
                                break;
                            }
                            else
                            {
                                Debug.WriteLine($"    没有找到名称为'{dataDictionaryName}'的DataDictionary");
                            }
                        }
                        else
                        {
                            Debug.WriteLine("    无法从过滤条件中提取DataDictionary名称");
                        }
                    }
                }
                else
                {
                    Debug.WriteLine("  CollectionSource没有过滤条件");
                }
            }
            else
            {
                Debug.WriteLine("  当前视图不是ListView或CollectionSource为null");
            }
            
            Debug.WriteLine("=== NewObjectController_ObjectCreated 执行完成 ===");
        }
    }
    
    /// <summary>
    /// 使用正则表达式从过滤条件字符串中提取DataDictionary名称
    /// </summary>
    /// <param name="criteriaString">过滤条件字符串</param>
    /// <returns>提取到的DataDictionary名称，失败则返回null</returns>
    private string ExtractDataDictionaryName(string criteriaString)
    {
        try
        {
            // 简化的正则表达式，避免复杂的转义
            string pattern = "DataDictionary\\.Name.*?['\"](.*?)['\"]";
            Match match = Regex.Match(criteriaString, pattern, RegexOptions.IgnoreCase);
            
            if (match.Success && match.Groups.Count > 1)
            {
                return match.Groups[1].Value;
            }
            
            return null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"    正则表达式提取失败: {ex.Message}");
            return null;
        }
    }

    protected override void OnDeactivated()
    {
        if (_newObjectController != null)
        {
            _newObjectController.ObjectCreated -= NewObjectController_ObjectCreated;
            _newObjectController = null;
        }
        base.OnDeactivated();
    }
}