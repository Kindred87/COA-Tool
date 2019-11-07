using System;
using System.Collections.Generic;
using System.Text;

namespace CoA_Tool.Templates
{
    class CustomFilter
    {
        public enum FilterTypes { Unassigned, Whitelist, Blacklist}

        public FilterTypes FilterType;

        public Template.ContentItems ContentItem;

        public List<string> Criteria;

        public bool IsValidFilter
        {
            get
            {
                if (FilterType != FilterTypes.Unassigned && ContentItem != Template.ContentItems.Unassigned && Criteria.Count > 0)
                    return true;
                else
                    return false;
            }
        }
        public CustomFilter()
        {
            Criteria = new List<string>();
            FilterType = FilterTypes.Unassigned;
            ContentItem = Template.ContentItems.Unassigned;
        }
    }
}
