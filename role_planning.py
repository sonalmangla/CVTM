import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta,date
NEW_MEMBER_FLAG = 4



def assign_roles():

    current_date = pd.Timestamp.today()
    past_meetings_df = pd.read_excel("CVTM Role Planning.xlsx",sheet_name="Meeting Planner")
    member_df = pd.read_excel("CVTM Role Planning.xlsx",sheet_name="Member")
    role_df = pd.read_excel("CVTM Role Planning.xlsx",sheet_name="Role")
    # meeting_df = pd.read_excel("CVTM Role Planning.xlsx",sheet_name="Meeting")

    active_member_df = member_df[member_df["active_to"].isnull()]

    active_member_df["active_since_days"] = (current_date - active_member_df["active_from"]).dt.days
    new_member_df = active_member_df[active_member_df["active_since_days"] < 28]
    old_members_df = active_member_df[active_member_df["active_since_days"] >= 28]
    active_member_df['membership_age'] = np.where(active_member_df['active_since_days'] < 28, '0 New Member',
                                                  np.where(active_member_df['active_since_days'] >= 28, '1 Senior', 'Not Known'))

    active_member_df['membership_age'] = np.where(active_member_df['new_member_flag'] == 'Y', '0 New Member',
                                                  np.where(active_member_df['new_member_flag'] =='N', '1 Senior', 'Not Known'))

    member_role_df = active_member_df.merge(role_df, how='cross')[["member_name","role_name"]]
    member_role_df = active_member_df.merge(member_role_df, how = "left", left_on=['member_name'], right_on=['member_name'])
    member_role_df = member_role_df.merge(role_df, how = "left", left_on=['role_name'], right_on=['role_name'])

    new_meeting_id = str(past_meetings_df['meeting_id'].max() + 1)
    new_meeting_date = (past_meetings_df['meeting_date'].max() + timedelta(days=7)).strftime("%d/%m/%Y")

    period_df = past_meetings_df.pivot_table(index='member_name', values='meeting_date', columns='role_name', aggfunc='max', fill_value='1990-01-01')

    role_last_performed_df = past_meetings_df.groupby(['member_name', 'role_name']).agg( role_last_performed_date = ('meeting_date','max'))
    all_members_roles_performed_df = member_role_df.merge(role_last_performed_df, how= "left", left_on=['member_name','role_name'], right_on=['member_name','role_name'])
    all_members_roles_performed_df.fillna('1990-01-01', inplace= True)
    # all_members_roles_performed_df['role_last_performed_date'] = pd.to_datetime(all_members_roles_performed_df['role_last_performed_date'])
    # all_members_roles_performed_df["days_since_event"] = (current_date - all_members_roles_performed_df["role_last_performed_date"]).dt.days

    all_members_roles_performed_df['rank'] = all_members_roles_performed_df.groupby('role_name')['role_last_performed_date'].rank(method='dense', ascending=True)
    all_members_roles_performed_sorted_df = all_members_roles_performed_df.sort_values(['role_level','role_last_performed_date','membership_age'],ascending=[True, True,True])
    members_shortlisted_df = all_members_roles_performed_sorted_df[all_members_roles_performed_sorted_df["rank"] == 1 ]
    # print(members_shortlisted_df)
    all_members_roles_performed_sorted_df[['member_name','role_name','role_last_performed_date']].to_csv("role_count.csv")

    role_array = role_df.sort_values(['role_level'], ascending= [True])['role_name'].to_numpy()
    members_role_assigned_dict = dict()

    for role in role_array:

        members_each_role_df = members_shortlisted_df[members_shortlisted_df['role_name'] == role]
        members_each_role_df = members_each_role_df.drop(members_each_role_df[members_each_role_df["member_name"].isin(list(members_role_assigned_dict.values()))].index.tolist())
        # row_num = random.randint(0,len(members_each_role_df)-1)
        row_num = 0
        print(row_num)
        member_role_assigned = members_each_role_df['member_name'].iloc[row_num]
        if role == "Pathways Speech" or role == "Evaluation":
            member_role_assigned_2 = members_each_role_df['member_name'].iloc[1]
            members_role_assigned_dict[role+'_2'] = member_role_assigned_2
            member_role_assigned_3 = members_each_role_df['member_name'].iloc[2]
            members_role_assigned_dict[role+'_3'] = member_role_assigned_3
        members_role_assigned_dict[role] = member_role_assigned




    # print(members_role_assigned_dict)

    members_role_assigned_df = pd.DataFrame.from_dict(members_role_assigned_dict,orient='index')
    members_role_assigned_df['meeting_date'] = new_meeting_date
    members_role_assigned_df['meeting_id'] = new_meeting_id
    members_role_assigned_df = members_role_assigned_df.reset_index()
    members_role_assigned_df['role_name'] = (members_role_assigned_df['index'].str.replace('_2', '')).str.replace('_3', '')
    members_role_assigned_df['member_name'] = members_role_assigned_df[0]
    # members_role_assigned_df.rename(columns={'meeting_date': 'meeting_date','meeting_id':'meeting_id','index':'role_name','0':'member_name'}, inplace=True)
    members_role_assigned_df[["meeting_date","meeting_id","role_name","member_name"]].to_csv("role_assigned.csv",index=False, encoding='utf-8')
    print("CSV file exported successfully!")



def extract_legacy_data():
    import pandas as pd

    # Read the workbook without using any row as a header so that the original row numbers are retained
    df = pd.read_excel('CVTM Role Planning - Legacy.xlsx', header=None, engine='openpyxl')

    # Identify every meeting column – those where row 2 and row 3 both contain a non-null entry
    meeting_cols = [c for c in df.columns[1:-1] if pd.notnull(df.iloc[1, c]) and pd.notnull(df.iloc[2, c])]

    # Build a list of dictionaries, one per person per meeting
    records = []
    for c in meeting_cols:
        meet_id = df.iloc[1, c]    # Meeting ID from row 2
        meet_date = df.iloc[2, c]  # Meeting Date from row 3
        for r in range(3, df.shape[0]):  # Loop from row 4 onward
            # Pick the name from column A if present, else fallback to column Q
            name = None
            if pd.notnull(df.iloc[r, 0]):
                name = df.iloc[r, 0]
            elif pd.notnull(df.iloc[r, df.shape[1] - 1]):
                name = df.iloc[r, df.shape[1] - 1]
            if name is not None:
                role = df.iloc[r, c]  # Role from the intersecting cell
                records.append({
                    'Meeting Date': meet_date,
                    'Meeting ID': meet_id,
                    'Member Name': name,
                    'Role': role
                })

    # Convert to a tidy DataFrame
    tidy_df = pd.DataFrame(records)

    # Optional: Save to CSV or Excel
    tidy_df.to_csv('extracted_roles.csv', index=False)
    # tidy_df.to_excel('extracted_roles.xlsx', index=False)



if __name__ == "__main__":
    extract_legacy_data()
    assign_roles()

    # all_members_roles_performed_df['rank'] = all_members_roles_performed_df.sort_values(['days_since_event'], \
    #                                  ascending=[False]) \
    #                       .groupby(['role_name']) \
    #                       .cumcount() + 1
