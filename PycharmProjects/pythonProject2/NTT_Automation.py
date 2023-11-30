import pandas as pd
import xlsxwriter

file_path = "New logs1.xlsx"

df = pd.read_excel(file_path)

pd.set_option('display.max_columns', None)

accurate_deboarding_count = 0
inaccurate_deboarding_count = 0
accurate_walking_count = 0
inaccurate_walking_count = 0

accurate_deboarding_log_ids = []
inaccurate_deboarding_log_ids = []
accurate_walking_log_ids = []
inaccurate_walking_log_ids = []

output_data = []

with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    for log_id, group_df in df.groupby('logId'):
        print(f"logId: {log_id}")
        if group_df.isnull().values.any():
            applied_rule = "Incomplete input data"
            print(f"Incomplete input data for logId: {log_id}")
        else:
            condition_applied = False
            applied_rule = "No business rule applied"

        individual_deboarding_times = []

        for index, row in group_df.iterrows():
            if row['inbound_de_board_type'] in ('Bus', 'Empty or unknown'):
                final_deboarding_time = row['cts_deboarding']
            elif row['inbound_de_board_type'] == 'Bridge' and row['no_rows_main_deck'] != 0:
                deboarding_time_per_row = row['cts_deboarding'] / (row['no_rows_main_deck'] / 2)
                final_deboarding_time = (deboarding_time_per_row * row['inbound_seatNumber_separated'] + 2) #(winter)

                # Check if subfleet is 388 and adjust the final deboarding time
                if row['inbound_subfleet'] == 388:
                    final_deboarding_time += 2  # Add 2 minutes for subfleet 388
                    applied_rule = "Additional 2 minutes for subfleet 388"
                else:
                    applied_rule = "No business rule applied"
            else:
                final_deboarding_time = None

            if final_deboarding_time is not None:
                individual_deboarding_times.append(final_deboarding_time)

        avg_deboarding_time = sum(individual_deboarding_times) / len(
            individual_deboarding_times) if individual_deboarding_times else None

        # Check if the                    applied_rule = "No business rule applied"
        #             else:
        #                 final_deboarding_time = None
        #
        #             if final_deboarding_time is not None:
        #                 individual_deboarding_times.append(final_deboarding_time)
        #
        #         avg_deboarding_time = sum(individual_deboarding_times) / len(
        #             individual_deboarding_times) if individual_deboarding_times else None
        #
        #         # Check if the 'ssr' column contains any of the specified values
        #         if group_df['ssr'].isin(['WCHR', 'WCHS', 'BLND', 'MEDA', 'STRC', 'UMNR']).any():
        #             # Set the total_walking_time to 100% of the original walking time
        #             group_df['total_walking_time'] = group_df['cts_walk_time'] - group_df['cts_checks']
        #             applied_rule = "Take 100% of walking time when ssr is WCHR,'WCHS', BLND, MEDA, STRC, or UMNR" 'ssr' column contains any of the specified values
        if group_df['ssr'].isin(['WCHR', 'WCHS', 'BLND', 'MEDA', 'STRC', 'UMNR']).any():
            # Set the total_walking_time to 100% of the original walking time
            group_df['total_walking_time'] = group_df['cts_walk_time'] - group_df['cts_checks']
            applied_rule = "Take 100% of walking time when ssr is WCHR,'WCHS', BLND, MEDA, STRC, or UMNR"

        if not applied_rule.startswith("Take 100% of walking time"):
            age_factor_condition = not any(group_df['tierLevel'].isin(['FTL', 'SEN', 'Gold', 'HON', 'HOC', 'F-Cl']))
            if age_factor_condition:
                if ((group_df['age'] >= 0) & (group_df['age'] <= 13)).any():
                    group_df.loc[(group_df['age'] >= 0) & (group_df['age'] <= 13), 'factor'] = 1.1
                    applied_rule = "age group 0 - 13 years with factor 1.1"

                # Determine the Age Group and Calculate Factor Based on Total Walking Time
                age_condition = (group_df['age'] >= 14) & (group_df['age'] <= 40)

                if age_condition.any():
                    group_df.loc[(group_df['age'] >= 14) & (group_df['age'] <= 40), 'factor'] = 0.8
                    applied_rule = "Reduction of walking times to 80% for age group 14-40"

                if ((group_df['age'] >= 41) & (group_df['age'] <= 60)).any():
                    group_df.loc[(group_df['age'] >= 41) & (group_df['age'] <= 60), 'factor'] = 1.0
                    applied_rule = " No change of walking time for age group 41-60 or age unknown"

                age_condition =(group_df['age'] > 60)
                if age_condition.any():
                    final_walking_time = group_df.iloc[0]['ntt_walking_time']
                    applied_rule = " Increase of walking times to 110% for age group > 60"
                    group_df['total_walking_time'] = final_walking_time

                if group_df['age'].isnull().any():
                    group_df.loc[group_df['age'].isnull(), 'factor'] = 1.0
                    applied_rule = " No change of walking time for age group 41-60 or age unknown"
            else:
                applied_rule = "No business rule applied"
                group_df['factor'] = 1.0

        if 'total_walking_time' in group_df.columns:
            avg_walking_time = group_df['total_walking_time'].mean()
        else:
            avg_walking_time = (group_df['cts_walk_time'] - group_df['cts_checks']).mean()

        final_walking_time = round(avg_walking_time + group_df.iloc[0]['cts_checks'], 2)

        if len(group_df) >= 15 and not any(group_df['tierLevel'].isin(['HON', 'SEN', 'Gold', 'FTL'])):
            applied_rule = "Increase of walking time to 125% for groups/linked pax >= 15"
            if 'total_walking_time' in group_df.columns:
                group_df['total_walking_time'] *= 1.25
            else:
                group_df['cts_walk_time'] = (group_df['cts_walk_time'] - group_df['cts_checks']) * 1.25 + group_df[
                    'cts_checks']
        else:
            # The new condition for HON, SEN, Gold, FTL with >=15 passengers
            hon_sen_gold_ftl_condition = group_df['tierLevel'].isin(['HON', 'SEN', 'Gold', 'FTL'])
            if len(group_df) >= 15 and hon_sen_gold_ftl_condition.any():
                applied_rule = "No change of walking time for groups/linked pax >= 15 and HON, SEN, Gold, FTL in the group"
                # Set the final walking time as the cts walking time
                group_df['final_walking_time'] = group_df['cts_walk_time']

        reduction_factor = 1.0  # Default reduction factor (no reduction)

        hon_hoc_fcl_condition = group_df['tierLevel'].isin(['HON', 'HOC', 'F-Cl'])
        if hon_hoc_fcl_condition.any():
            applied_rule = "Reduction of walking time to 60% for HON, HOC, F-Cl"
            if 'total_walking_time' in group_df.columns:
                group_df.loc[hon_hoc_fcl_condition, 'total_walking_time'] *= 0.6

            else:
                group_df.loc[hon_hoc_fcl_condition, 'cts_walk_time'] = (group_df.loc[hon_hoc_fcl_condition, 'cts_walk_time'] - group_df.loc[hon_hoc_fcl_condition, 'cts_checks']) * 0.6 + group_df.loc[hon_hoc_fcl_condition, 'cts_checks']
            condition_applied = True

        # New Business Rule: Reduction of walking time to 75% for FTL, SEN, Gold
        ftl_sen_gold_condition = group_df['tierLevel'].isin(['FTL', 'SEN', 'Gold'])
        if ftl_sen_gold_condition.any():
            same_group_ids = not group_df['group_id'].nunique() == 1
            separate_group_ids = group_df['group_id'].nunique() == len(group_df)
            has_infant_child_elderly = group_df['age'].apply(lambda x: x < 14 or x > 60)


            if group_df['group_id'].empty or ('NA' in group_df['group_id'].values and separate_group_ids):
                # Check if all passengers in the group are FTL, SEN, or Gold
                all_ftl_sen_gold = group_df['tierLevel'].isin(['FTL', 'SEN', 'Gold']).all()

                if all_ftl_sen_gold :
                    reduction_factor = 0.75  # 75% reduction
                    applied_rule = "Reduction of walking time to 75% for FTL, SEN, Gold"
                else:
                    reduction_factor = 1.0  # Default reduction factor (no reduction)
                    applied_rule = "No business rule applied"

            elif same_group_ids or (separate_group_ids and has_infant_child_elderly.any()):
                reduction_factor = 0.9 # 90% reductions
                group_df['final_walking_time'] = group_df['ntt_walking_time']
                applied_rule = "Reduction of walking time to 90% for FTL, SEN, Gold"

            else:
                reduction_factor = 0.8  # 80% reductions
                applied_rule = "Reduction of walking time to 80% for FTL, SEN, Gold"

            if 'total_walking_time' in group_df.columns:
                group_df['total_walking_time'] *= reduction_factor
            else:
                group_df['cts_walk_time'] = (group_df['cts_walk_time'] - group_df['cts_checks']) * reduction_factor + group_df['cts_checks']
            condition_applied = True

        if 'total_walking_time' in group_df.columns:
            avg_walking_time = group_df['total_walking_time'].mean()
        else:
            avg_walking_time = (group_df['cts_walk_time'] - group_df['cts_checks']).mean()

        final_walking_time = round(avg_walking_time + group_df.iloc[0]['cts_checks'], 2)

        print(group_df[
                  ['logId', 'group_id', 'uniqueCustomerId', 'cts_deboarding', 'no_rows_main_deck', 'age', 'cts_checks',
                   'cts_walk_time', 'ntt_de-boarding-time', 'ntt_walking_time', 'inbound_de_board_type', 'ssr',
                   'tierLevel', 'inbound_subfleet', 'inbound_seatNumber', 'inbound_seatNumber_separated']].to_string(index=False))

        if avg_deboarding_time is not None:
            avg_deboarding_time = round(avg_deboarding_time, 2)
            ntt_deboarding_time = round(group_df.iloc[0]['ntt_de-boarding-time'], 2)
            print("Final Deboarding Time:", avg_deboarding_time)
            if avg_deboarding_time == ntt_deboarding_time:
                print("Deboarding Time is accurate")
                accurate_deboarding_count += 1
                accurate_deboarding_log_ids.append(log_id)
            else:
                print("Deboarding Time doesn't match")
                inaccurate_deboarding_count += 1
        else:
            ntt_deboarding_time = None

        if avg_walking_time is not None:
            avg_walking_time = round(avg_walking_time, 2)
            ntt_walking_time = round(group_df.iloc[0]['ntt_walking_time'], 2)
            print("Final Walking Time:", final_walking_time)

            if final_walking_time == ntt_walking_time:
                walking_time_status = "Accurate"  # Walking time is accurate
            else:
                walking_time_status = "Doesn't match"  # Walking time doesn't match

            print("Walking Time is", walking_time_status)

            if walking_time_status == "Accurate":
                accurate_walking_count += 1
                accurate_walking_log_ids.append(log_id)
            else:
                inaccurate_walking_count += 1
                inaccurate_walking_log_ids.append(log_id)  # Add the inaccurate walking log ID to the list
        else:
            walking_time_status = "N/A"

        log_details = {
            'Logid': log_id,
            'ntt de-boarding-time': ntt_deboarding_time,
            'ntt walking-time': ntt_walking_time,
            'Final Deboarding time': avg_deboarding_time if avg_deboarding_time is not None else 'N/A',
            'Final Walking time': final_walking_time,
            'Deboarding time status': "Accurate" if avg_deboarding_time == ntt_deboarding_time else "Doesn't match",
            'Walking time status': walking_time_status,  # Use the walking_time_status variable
            'Business rule used': applied_rule
        }

        output_data.append(log_details)

    print(f"Accurate Deboarding Time Logs Count: {accurate_deboarding_count}")
    print(f"Inaccurate Deboarding Time Logs Count: {inaccurate_deboarding_count}")
    print(f"Accurate Walking Time Logs Count: {accurate_walking_count}")
    print(f"Inaccurate Walking Time Logs Count: {inaccurate_walking_count}")  # Print inaccurate walking time count

    # Print log IDs for accurate and inaccurate walking times
    print("Accurate Deboarding Log IDs:")
    for log_id in accurate_deboarding_log_ids:
        print(log_id)
    print("Inaccurate Deboarding Log IDs:")
    for log_id in inaccurate_deboarding_log_ids:
        print(log_id)
    print("Accurate Walking Log IDs:")
    for log_id in accurate_walking_log_ids:
        print(log_id)
    print("Inaccurate Walking Log IDs:")  # Print inaccurate walking log IDs
    for log_id in inaccurate_walking_log_ids:
        print(log_id)

    output_df = pd.DataFrame(output_data)

    output_df.to_excel(writer, sheet_name='Log_Details', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Log_Details']

    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'align': 'center',
        'fg_color': '#D7E4BC',
        'border': 1
    })

    for i, col in enumerate(output_df.columns):
        column_len = max(output_df[col].astype(str).str.len().max(), len(col)) + 2
        worksheet.set_column(i, i, column_len)
        worksheet.write(0, i, col, header_format)

    summary_data = {
        'Accurate Deboarding Time Logs Count': [accurate_deboarding_count],
        'Inaccurate Deboarding Time Logs Count': [inaccurate_deboarding_count],
        'Accurate Walking Time Logs Count': [accurate_walking_count],
        'Inaccurate Walking Time Logs Count': [inaccurate_walking_count],
    }

    summary_df = pd.DataFrame(summary_data)

    summary_df.to_excel(writer, sheet_name='Summary', index=False)

    summary_worksheet = writer.sheets['Summary']

    for i, col in enumerate(summary_df.columns):
        column_len = max(summary_df[col].astype(str).str.len().max(), len(col)) + 2
        summary_worksheet.set_column(i, i, column_len)
        summary_worksheet.write(0, i, col, header_format)
