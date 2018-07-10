namespace XmlToSql
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Area")]
    public partial class Area
    {
        public int AreaId { get; set; }

        public int? TypeArea { get; set; }

        [StringLength(500)]
        public string Name { get; set; }

        public int? ParentId { get; set; }

        public bool? NoSplit { get; set; }

        public bool? AssemblyArea { get; set; }

        [StringLength(500)]
        public string FullName { get; set; }

        public bool? MultipleOrders { get; set; }

        public bool? HiddenArea { get; set; }

        [StringLength(5000)]
        public string IP { get; set; }

        public int PavilionId { get; set; }

        public int TypeId { get; set; }

        public int? OrderExecution { get; set; }

        public int? Dependence { get; set; }

        public int? WorkingPeople { get; set; }

        public int? ComponentTypeId { get; set; }

        public int? GroupId { get; set; }

        public int? Segment { get; set; }
    }
}
