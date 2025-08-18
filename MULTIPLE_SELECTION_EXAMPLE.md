# Multiple Selection Example

## How Multiple Filters Work

Your enhanced lead generation tool now supports multiple selections just like your n8n workflow.

### Example Selection:

**Job Titles:** Founder, CEO, C suite  
**Company Sizes:** 1-10, 11-20, 21-50

### Generated Apollo URL:
```
https://app.apollo.io/#/people?page=1
&contactEmailStatusV2[]=verified
&existFields[]=person_title_normalized
&existFields[]=organization_domain
&personLocations[]=Singapore
&personLocations[]=Singapore%2C%20Singapore
&sortAscending=true
&sortByField=sanitized_organization_name_unanalyzed
&personTitles[]=Founder
&personTitles[]=CEO
&personTitles[]=C%20suite
&organizationNumEmployeesRanges[]=1%2C10
&organizationNumEmployeesRanges[]=11%2C20
&organizationNumEmployeesRanges[]=21%2C50
```

### What This Finds:
- **Job Titles**: People with ANY of the selected titles (Founder OR CEO OR C suite)
- **Company Sizes**: Companies with ANY of the selected sizes (1-10 OR 11-20 OR 21-50 employees)
- **Location**: Singapore (fixed filter)
- **Email**: Only verified emails
- **Result**: All combinations that match the criteria

### Example Results:
- John Doe, **CEO** at TechCorp (**15 employees**) ✅
- Jane Smith, **Founder** at StartupABC (**8 employees**) ✅ 
- Mike Johnson, **C suite** at GrowthCo (**35 employees**) ✅
- Sarah Lee, **Director** at BigCorp (**5 employees**) ❌ (wrong title)
- Tom Wilson, **CEO** at Enterprise (**500 employees**) ❌ (wrong size)

### UI Features:
1. **Selection Tags**: Visual tags show what you've selected
2. **Applied Filters**: Results show exactly which filters were used
3. **Multiple Validation**: Ensures at least one selection in each category
4. **Progress Tracking**: Shows how many leads match your combined criteria

This matches your n8n workflow behavior exactly - multiple selections create OR conditions within each filter type, and AND conditions between filter types.