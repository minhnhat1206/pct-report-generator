import pandas as pd
import re
import math
import logging

def parse_study_time(time_str):
    """
    Parses 'HH:MM' or 'HH:MM:SS' into total minutes.
    Matches logic in report_grade_10.py.
    Returns 0 if parsing fails.
    """
    if pd.isna(time_str):
        return 0
    
    # Handle already numeric values (e.g. from Excel calculation)
    if isinstance(time_str, (int, float)):
        return float(time_str)

    if not isinstance(time_str, str):
        return 0
    
    try:
        parts = list(map(int, time_str.split(':')))
        if len(parts) >= 2:
            hours, minutes = parts[0], parts[1]
            return hours * 60 + minutes
    except Exception:
        pass
        
    return 0

def safe_float(value, default=0.0):
    """Safely converts value to float, handling None, NaN, Infinity, strings."""
    try:
        if pd.isna(value):
            return default
        val = float(value)
        if math.isinf(val):
            return default
        return val
    except (ValueError, TypeError):
        return default

def calculate_stats(data):
    """
    Calculates statistics for the dashboard from the processed dataframe.
    """
    if data is None or data.empty:
        return None

    stats = {
        'class_stats': [],
        'status_counts': {}
    }

    try:
        # Ensure Study_Time is parsed to minutes for aggregation
        # We work on a copy to avoid affecting the main report generation if it uses the original string
        df_analysis = data.copy()
        
        # Check if 'Study_Time_Minutes' already exists (from report logic) or parse it
        if 'Study_Time_Minutes' not in df_analysis.columns:
            # Handle case where Study_Time might be missing
            if 'Study_Time' in df_analysis.columns:
                df_analysis['Study_Time_Minutes'] = df_analysis['Study_Time'].apply(parse_study_time)
            else:
                 df_analysis['Study_Time_Minutes'] = 0

        # FORCE NUMERIC CONVERSION for critical columns
        # invalid parsing will be set as NaN
        if 'Unitslessons_Passed' in df_analysis.columns:
             df_analysis['Unitslessons_Passed'] = pd.to_numeric(df_analysis['Unitslessons_Passed'], errors='coerce')
        
        if 'Unitslessons_Studied' in df_analysis.columns:
             df_analysis['Unitslessons_Studied'] = pd.to_numeric(df_analysis['Unitslessons_Studied'], errors='coerce')

        # Group by Class
        # 'English_Class_y' is the standard column name from report_grade_10/11 clean_data
        if 'English_Class_y' in df_analysis.columns:
            group_col = 'English_Class_y'
        elif 'English Class' in df_analysis.columns: # Fallback
            group_col = 'English Class'
        else:
            logging.warning("Column 'English_Class_y' not found in data for analysis.")
            return None

        # Calculate averages per class
        class_groups = df_analysis.groupby(group_col)
        
        for class_name, group in class_groups:
            # Safe aggregation
            avg_progress = group['Unitslessons_Passed'].mean() if 'Unitslessons_Passed' in group.columns else 0
            avg_studied = group['Unitslessons_Studied'].mean() if 'Unitslessons_Studied' in group.columns else 0
            
            # New Time Logic
            total_minutes = group['Study_Time_Minutes'].sum() if 'Study_Time_Minutes' in group.columns else 0
            total_passed = group['Unitslessons_Passed'].sum() if 'Unitslessons_Passed' in group.columns else 0
            total_studied = group['Unitslessons_Studied'].sum() if 'Unitslessons_Studied' in group.columns else 0
            
            avg_total_time = group['Study_Time_Minutes'].mean() if 'Study_Time_Minutes' in group.columns else 0 # Phút / Học sinh
            
            # Avoid division by zero
            if total_studied > 0:
                avg_time_per_studied = total_minutes / total_studied
            else:
                avg_time_per_studied = 0
            
            # Status Counts per class
            on_track_count = 0
            behind_count = 0
            if 'Status' in group.columns:
                status_counts = group['Status'].value_counts()
                # 'keep up' and 'far away' are considered on track/ahead
                on_track_count = status_counts.get('keep up', 0) + status_counts.get('far away', 0)
                behind_count = status_counts.get('late', 0)

            stats['class_stats'].append({
                'className': str(class_name),
                'avgProgress': round(safe_float(avg_progress), 1),
                'avgStudied': round(safe_float(avg_studied), 1),
                'avgTotalTime': round(safe_float(avg_total_time), 1),
                'avgTimePerStudied': round(safe_float(avg_time_per_studied), 1),
                'onTrack': int(on_track_count),
                'behind': int(behind_count)
            })

        # Global Status Counts - Map to Vietnamese
        if 'Status' in df_analysis.columns:
            counts = df_analysis['Status'].value_counts().to_dict()
            stats['status_counts'] = {
                'vượt kế hoạch': int(counts.get('far away', 0)),
                'đúng kế hoạch': int(counts.get('keep up', 0)),
                'chậm hơn kế hoạch': int(counts.get('late', 0))
            }

    except Exception as e:
        logging.error(f"Error in calculate_stats: {e}")
        # Return what we have or None? 
        # Better to return None to avoid partial/misleading data unless we are sure
        return None

    return stats
