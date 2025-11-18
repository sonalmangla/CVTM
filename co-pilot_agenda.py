import pandas as pd

# Load data from Excel
file_path  = 'CVTM Role Planning.xlsx'
meeting_df = pd.read_excel(file_path, sheet_name='Meeting Planner', engine='openpyxl')
member_df  = pd.read_excel(file_path, sheet_name='Member', engine='openpyxl')
role_df    = pd.read_excel(file_path, sheet_name='Role', engine='openpyxl')

def generate_meeting_agenda(meeting_df, member_df, role_df, apologies=None, next_meeting_date=None):
    """
    Generate the next meeting agenda based on business rules:
    - Fair rotation of roles
    - New members get roles they haven't done yet
    - Roles for new members assigned from easiest to hardest (role_level)
    - Apologies excluded
    - Meeting date defaults to one week after last meeting
    """
    active = member_df[member_df['active_flag'] == 'Y']['member_name'].tolist()
    apologies = apologies or []
    available = [m for m in active if m not in apologies]

    # Sort roles by complexity level (easiest first)
    sorted_roles = role_df.sort_values('role_level')['role_name'].tolist()

    # Historical counts of member-role assignments
    counts = meeting_df.pivot_table(index='member_name', columns='role_name', aggfunc='size', fill_value=0)
    for m in active:
        if m not in counts.index:
            counts.loc[m] = 0
    for r in sorted_roles:
        if r not in counts.columns:
            counts[r] = 0
    counts = counts.sort_index().sort_index(axis=1)

    # Identify new members and track roles they've done
    new_members = member_df[member_df['new_member_flag'] == 'Y']['member_name'].tolist()
    roles_done = {m: set(meeting_df[meeting_df['member_name'] == m]['role_name'].tolist()) for m in new_members}
    for m in new_members:
        roles_done.setdefault(m, set())

    assigned = {}
    assigned_members = set()

    # Determine next meeting date
    if next_meeting_date is None:
        last_date = meeting_df['meeting_date'].max()
        next_meeting_date = pd.to_datetime(last_date) + pd.DateOffset(weeks=1)

    # Assign roles
    for role in sorted_roles:
        elig = [m for m in available if m not in assigned_members]
        if not elig:
            assigned_members = set()
            elig = [m for m in available if m not in assigned_members]

        # Prioritise new members who haven't done this role
        elig_new = [m for m in elig if m in new_members and role not in roles_done.get(m, set())]
        if elig_new:
            elig_new_sorted = sorted(elig_new, key=lambda x: (len(roles_done.get(x, set())), x))
            chosen = elig_new_sorted[0]
        else:
            min_count = min(counts.loc[m, role] for m in elig)
            elig_min = [m for m in elig if counts.loc[m, role] == min_count]
            elig_min = sorted(elig_min)
            chosen = elig_min[0] if elig_min else None

        assigned[role] = chosen
        if chosen is not None:
            assigned_members.add(chosen)

    agenda = pd.DataFrame({
        'meeting_date': next_meeting_date,
        'role_name': list(assigned.keys()),
        'member_name': list(assigned.values())
    })
    return agenda

# Example usage
next_agenda = generate_meeting_agenda(meeting_df, member_df, role_df)
print(next_agenda)
