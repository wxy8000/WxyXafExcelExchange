using DevExpress.ExpressApp.Model;
using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.Persistent.Validation;
using DevExpress.Xpo;
using Wxy.Xaf.DataDictionary;
using Wxy.Xaf.RememberLast.Attributes;
using Wxy.Xaf.ExcelExchange;


namespace XafExcelDemo.Module.BusinessObjects
{
    [DefaultClassOptions]
    [ExcelImportExport]
    public class 员工 : BaseObject
    {
        public 员工(Session session)
            : base(session)
        {
        }

        public override void AfterConstruction()
        {
            base.AfterConstruction();
            // 设置默认值
            // 注意：这些值可能会被 RememberLast 控制器覆盖
            // RememberLast 在对象视图激活后执行
            入职日期 = DateTime.Now;
            是否在职 = true;
        }

        private string _工号;
        private string _姓名;
        private DataDictionaryItem _性别;
        private DataDictionaryItem _民族;
        private DataDictionaryItem _学历;
        private DataDictionaryItem _政治面貌;
        private DataDictionaryItem _员工状态;
        private DataDictionaryItem _部门;
        private string _职位;
        private DateTime _入职日期;
        private string _邮箱;
        private string _手机;
        private bool _是否在职;
        private string _备注;

        [Size(50)]
        [RuleRequiredField]
        [RuleUniqueValue]
        [RememberLast]
        public string 工号
        {
            get => _工号;
            set => SetPropertyValue(nameof(工号), ref _工号, value);
        }

        [Size(SizeAttribute.DefaultStringMappingFieldSize)]
        [RuleRequiredField]
        [RememberLast]
        public string 姓名
        {
            get => _姓名;
            set => SetPropertyValue(nameof(姓名), ref _姓名, value);
        }

        [DataDictionary("性别", "男|1;女|2;未知|0")]
        [RememberLast]
        public DataDictionaryItem 性别
        {
            get => _性别;
            set => SetPropertyValue(nameof(性别), ref _性别, value);
        }

        [DataDictionary("民族", "汉族|1;壮族|2;回族|3;满族|4;其他|99")]
        [RememberLast]
        public DataDictionaryItem 民族
        {
            get => _民族;
            set => SetPropertyValue(nameof(民族), ref _民族, value);
        }

        [DataDictionary("学历", "博士|1;硕士|2;本科|3;大专|4;高中|5;初中|6;小学|7;其他|9")]
        [RememberLast]
        public DataDictionaryItem 学历
        {
            get => _学历;
            set => SetPropertyValue(nameof(学历), ref _学历, value);
        }

        [DataDictionary("政治面貌", "党员|1;团员|2;群众|3;民主党派|4")]
        [RememberLast]
        public DataDictionaryItem 政治面貌
        {
            get => _政治面貌;
            set => SetPropertyValue(nameof(政治面貌), ref _政治面貌, value);
        }

        [DataDictionary("员工状态", "在职|1;试用期|2;离职|3;退休|4")]
        [RememberLast]
        public DataDictionaryItem 员工状态
        {
            get => _员工状态;
            set => SetPropertyValue(nameof(员工状态), ref _员工状态, value);
        }

        [DataDictionary("部门类型", "技术部|1;市场部|2;人事部|3;财务部|4;行政部|5")]
        [RememberLast]
        public DataDictionaryItem 部门
        {
            get => _部门;
            set => SetPropertyValue(nameof(部门), ref _部门, value);
        }

        [Size(100)]
        [RememberLast]
        public string 职位
        {
            get => _职位;
            set => SetPropertyValue(nameof(职位), ref _职位, value);
        }

        [RememberLast]
        public DateTime 入职日期
        {
            get => _入职日期;
            set => SetPropertyValue(nameof(入职日期), ref _入职日期, value);
        }

        [Size(200)]
        [RememberLast]
        public string 邮箱
        {
            get => _邮箱;
            set => SetPropertyValue(nameof(邮箱), ref _邮箱, value);
        }

        [Size(50)]

        public string 手机
        {
            get => _手机;
            set => SetPropertyValue(nameof(手机), ref _手机, value);
        }

        [RememberLast]
        public bool 是否在职
        {
            get => _是否在职;
            set => SetPropertyValue(nameof(是否在职), ref _是否在职, value);
        }

        [Size(SizeAttribute.Unlimited)]
        [RememberLast]
        public string 备注
        {
            get => _备注;
            set => SetPropertyValue(nameof(备注), ref _备注, value);
        }
    }
}