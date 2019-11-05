using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Templates
{
    class CustomFilter
    {
        public enum FilterType { In, Out}

        public FilterType Filter;

        public Template.ContentItems ContentItem;

        public List<string> Criteria;
        public CustomFilter()
        {
            Criteria = new List<string>();
        }
    }
}
