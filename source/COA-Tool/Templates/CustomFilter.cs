using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Templates
{
    class CustomFilter
    {
        public enum FilterType { Unassigned, Whitelist, Blacklist}

        public FilterType Filter;

        public Template.ContentItems ContentItem;

        public List<string> Criteria;

        public bool ValidFilter
        {
            get
            {
                if (Filter != FilterType.Unassigned && ContentItem != Template.ContentItems.Unassigned && Criteria.Count > 0)
                    return true;
                else
                    return false;
            }
        }
        public CustomFilter()
        {
            Criteria = new List<string>();
            Filter = FilterType.Unassigned;
            ContentItem = Template.ContentItems.Unassigned;
        }
    }
}
