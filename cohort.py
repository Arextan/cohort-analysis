import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_excel('invoices.xlsx')

# customer frequency
df['Records'] = df.groupby('Customer Number')['Customer Number'].transform('count')

# customer first & last date
df['First Date'] = df.groupby('Customer Number')['Year'].transform('min')
df['Last Date'] = df.groupby('Customer Number')['Year'].transform('max')

# last annual sale amount to each customer
last_value = df[['Customer Number', 'Year', 'Last Date', 'Document Amount']]
last_value = last_value[last_value['Year'] == last_value['Last Date']]
last_value['Last Value'] = \
    last_value.groupby(['Customer Number', 'Last Date'])['Document Amount'].transform('sum')
last_value = last_value.drop_duplicates(subset='Customer Number')
df_last_value = last_value[['Customer Number', 'Last Value']]

# summary statistics
sum_stats = df.groupby('Customer Number', as_index=False)['Document Amount'].agg(
    {'Total': 'sum', 'Average': 'mean', 'Median': 'median', 'Std Dev.': 'std'})
sum_stats = sum_stats.fillna(0)

# total sales for each year
year_counter = df['First Date'].min()
for years in range(df['Last Date'].max() - df['First Date'].min() + 1):
    df[str(year_counter)] = df[df['Year'] == year_counter].groupby(['Customer Number', 'Year'])[
        'Document Amount'].transform('sum').round(0)
    year_counter += 1

years = df[['Customer Number', 'Year', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009',
            '2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019']]
years = years.drop_duplicates(subset=['Customer Number', 'Year'])

# number of years servicing customer
df['Years'] = years.groupby('Customer Number')['Year'].transform('count')

years = years.groupby('Customer Number', sort=False, as_index=False).agg('sum')
df_years = years.drop(columns='Year')

# cleaned df with relevant columns
df_new = df[['Customer Number', 'Customer Name', 'Records', 'First Date', 'Last Date', 'Years']]
df_new = df_new.drop_duplicates(subset='Customer Number')
df_main = df_new.merge(sum_stats, on='Customer Number').merge(df_last_value, on='Customer Number') \
    .merge(df_years, on='Customer Number')
df_main.insert(7, 'Last Value', df_main.pop('Last Value'))
df_main.iloc[:, 5:] = df_main.iloc[:, 5:].round(0).astype('int64')

# customers cohort by vintage year
vintage_years = {('vintage_' + str(df_main['First Date'].min() + i)): df_main[
    df_main['First Date'] == df_main['First Date'].min() + i].drop(df.columns[11:11 + v], axis=1)
                 for i, v in enumerate(range((df_main['Last Date'].max() - df_main['First Date'].min() + 1)))}

# all customers in a year
all_years = {('all_' + str(df_main['First Date'].min() + i)): df_main[df_main[str(df_main['First Date'].min() + i)] > 0]
             for i, v in enumerate(range(df_main['Last Date'].max() - df_main['First Date'].min() + 1))}

# customer originations bar chart
vintage_sum = [vintage_years['vintage_' + str(2004 + i)][str(2004 + i)].sum() for i, v in
               enumerate(range(df_main['Last Date'].max() - 2004))]
vintage_count = [vintage_years['vintage_' + str(2004 + i)]['Customer Number'].count() for i, v in
                 enumerate(range(df_main['Last Date'].max() - 2004))]
vin_years = [2004 + i for i, v in enumerate(range(df_main['Last Date'].max() - 2004))]
cust_orig = {'New Customer Originations': vintage_sum, 'Customer Count': vintage_count, 'Years': vin_years}
df_cust_orig = pd.DataFrame(cust_orig)

ax_sum = df_cust_orig.plot.bar(x='Years', y='New Customer Originations', ylim=(0, 4000000), legend=False, rot=0)
ax_sum.ticklabel_format(axis='y', useOffset=False, style='plain')
ax_count = df_cust_orig.plot(x='Years', y='Customer Count', legend=False, color='r', ax=ax_sum, use_index=False,
                             secondary_y=True, mark_right=False)
ax_count.set_ylim(0, 250)
figure = plt.gcf()
figure.set_size_inches(12.8, 7.2)
plt.savefig('Customer Originations (unused).png', dpi=800)

# total invoices by year stacked area chart
first_date = df_main[['First Date', 'Customer Number']]
vin_totals = df_years.merge(first_date, on='Customer Number')
vin_totals = vin_totals.drop(columns=['Customer Number', '2002', '2019'])
vin_totals = vin_totals.groupby('First Date').agg('sum').astype('int64')
vin_totals = vin_totals.iloc[1:-1].T
area_years = [2003 + i for i, v in enumerate(range(16))]
vin_totals['Years'] = area_years

area_chart = vin_totals.plot.area(x='Years', xlabel="", colormap='tab20', title='Total Invoices by Vintage Year',
                                  linewidth=0, xlim=(2003, 2018), xticks=area_years)
area_chart.ticklabel_format(axis='y', useOffset=False, style='plain')
area_chart.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05), ncol=16, labelspacing=0, fontsize=6)
figure = plt.gcf()
figure.set_size_inches(12.8, 7.2)
plt.savefig('Stacked Area.png', dpi=800)

# cohort analysis line plot
total_inv = vin_totals.drop(columns='Years').T
total_years = {
    (str(z + 1)): [(total_inv.iloc[i, i + z] / total_inv.iloc[i, i] * 100) for i, v in
                   enumerate(range(16 - z))] for z, y in enumerate(range(16))}
df_total_years = pd.DataFrame.from_dict(total_years, orient='index')
df_total_years = df_total_years.T
avg_years = df_total_years.agg('mean').round(0).astype('int64')
excl_years = df_total_years.drop(index=[5, 6, 7, 8])
excl_recess = excl_years.agg('mean').round(0).astype('int64')
excl_years = excl_years.drop(index=0, columns='16')
excl_2003 = excl_years.agg('mean').round(0).astype('int64')
cohort_years = [i for i, v in enumerate(range(16))]

cohort_avg = pd.concat([avg_years, excl_recess, excl_2003], axis=1)
cohort_plot = cohort_avg.plot(xticks=cohort_years, ylim=(0, 100), legend=True, title='Cohort Analysis',
                              ylabel='Revenues as a Percentage of Year 0', xlabel='Years', color=['b', 'r', 'g'])
cohort_plot.legend(['Average', 'Excluding Recession', 'Excluding 2003 & Recession'], loc='upper center',
                   bbox_to_anchor=(0.5, -0.08), ncol=3)
figure = plt.gcf()
figure.set_size_inches(12.8, 7.2)
plt.savefig('Cohort Analysis.png', dpi=800)
# plt.show()

# export to excel
with pd.ExcelWriter('Python Cohort Analysis.xlsx') as writer:
    df_main.to_excel(writer, sheet_name='Values', index=False)
    for i, v in enumerate(range((df_main['Last Date'].max() - df_main['First Date'].min() + 1))):
        vintage_years['vintage_' + str(df_main['First Date'].min() + i)].to_excel(writer, sheet_name=str(
            df_main['First Date'].min() + i) + ' Vintage', index=False)
        all_years['all_' + str(df_main['First Date'].min() + i)].to_excel(writer, sheet_name=str(
            df_main['First Date'].min() + i) + ' All', index=False)
print('DataFrame written to Excel file successfully.')
