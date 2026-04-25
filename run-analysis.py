import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import argparse

def convert_pace_to_minutes(pace_str):
    """Convert pace string (e.g., '5:41') to minutes (e.g., 5.683)"""
    if pd.isna(pace_str):
        return None
    try:
        minutes, seconds = map(float, pace_str.split(':'))
        return minutes + seconds/60
    except:
        return None

def load_and_prepare_data(file_path):
    """Load and prepare the running data"""
    df = pd.read_excel(file_path, sheet_name='Garmin Activities')
    
    # Convert date column to datetime
    df['Date'] = pd.to_datetime(df['Date'])
    
    # Filter only running activities and exclude extreme distances
    running_df = df[
        (df['activityType'].str.contains('running', case=False, na=False)) & 
        (df['Distance (km)'] <= 700)
    ].copy()
    
    # Convert pace to numeric format (minutes per km)
    running_df['Pace (min/km)'] = running_df['Pace (min/km)'].apply(convert_pace_to_minutes)
    
    # Create month-year and year-month columns for analysis
    running_df['Month'] = running_df['Date'].dt.month
    running_df['Year'] = running_df['Date'].dt.year
    running_df['Month-Year'] = running_df['Date'].dt.to_period('M')
    running_df['Year-Month'] = running_df['Date'].dt.strftime('%Y-%m')
    
    return running_df

def plot_running_data(running_df, start_date=None, end_date=None, chart_type='bar', output_file=None):
    """Plot running distance data with optional date range filtering"""
    
    # Set default date range if not specified
    if start_date is None:
        start_date = running_df['Date'].min()
    if end_date is None:
        end_date = running_df['Date'].max()
    
    # Convert string dates to datetime if needed
    if isinstance(start_date, str):
        start_date = pd.to_datetime(start_date)
    if isinstance(end_date, str):
        end_date = pd.to_datetime(end_date)
    
    # Filter data by date range
    mask = (running_df['Date'] >= start_date) & (running_df['Date'] <= end_date)
    filtered_df = running_df.loc[mask]
    
    if len(filtered_df) == 0:
        print("No data available for selected date range")
        return
    
    # Group by month-year and calculate metrics
    monthly_stats = filtered_df.groupby('Year-Month').agg({
        'Distance (km)': ['sum', 'mean', 'count'],
        'Heart Rate (bpm)': 'mean',
        'Pace (min/km)': 'mean'
    }).reset_index()
    
    monthly_stats.columns = ['Year-Month', 'Total Distance (km)', 'Avg Distance (km)', 
                           'Number of Runs', 'Avg Heart Rate (bpm)', 'Avg Pace (min/km)']
    
    # Convert pace back to minutes:seconds format for display
    monthly_stats['Avg Pace (min/km)'] = monthly_stats['Avg Pace (min/km)'].apply(
        lambda x: f"{int(x)}:{int((x % 1) * 60):02d}" if pd.notna(x) else None
    )
    
    # Plot settings
    plt.figure(figsize=(12, 6))
    sns.set_style("whitegrid")
    title = f'Running Distance ({start_date.strftime("%Y-%m-%d")} to {end_date.strftime("%Y-%m-%d")})'
    plt.title(title, fontsize=16)
    plt.xlabel('Month', fontsize=12)
    plt.ylabel('Total Distance (km)', fontsize=12)
    
    if chart_type.lower() == 'bar':
        plot = sns.barplot(
            x='Year-Month', 
            y='Total Distance (km)', 
            data=monthly_stats,
            palette='viridis'
        )
        # Add value labels
        for p in plot.patches:
            plot.annotate(
                f"{p.get_height():.1f}",
                (p.get_x() + p.get_width() / 2., p.get_height()),
                ha='center', va='center',
                xytext=(0, 5),
                textcoords='offset points'
            )
    else:
        plot = sns.lineplot(
            x='Year-Month',
            y='Total Distance (km)',
            data=monthly_stats,
            marker='o',
            color='royalblue',
            linewidth=2.5
        )
        # Add data point labels
        for x, y in zip(range(len(monthly_stats)), monthly_stats['Total Distance (km)']):
            plt.text(
                x, y+5, f"{y:.1f}",
                color='black',
                ha='center',
                va='bottom'
            )
    
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    
    # Save or show the plot
    if output_file:
        plt.savefig(output_file)
        print(f"Plot saved to {output_file}")
    else:
        plt.show()
    
    # Print statistics table
    print("\nMonthly Running Statistics:")
    print(monthly_stats.to_string(index=False))

def main():
    # Set up command line arguments
    parser = argparse.ArgumentParser(description='Analyze running distance data')
    parser.add_argument('file_path', help='Path to the Excel file with running data')
    parser.add_argument('--start', help='Start date (YYYY-MM-DD)', default=None)
    parser.add_argument('--end', help='End date (YYYY-MM-DD)', default=None)
    parser.add_argument('--chart-type', choices=['bar', 'line'], default='bar',
                       help='Type of chart to display (bar or line)')
    parser.add_argument('--output', help='Output file to save the chart', default=None)
    
    args = parser.parse_args()
    
    try:
        # Load and prepare data
        running_df = load_and_prepare_data(args.file_path)
        
        # Plot the data
        plot_running_data(
            running_df,
            start_date=args.start,
            end_date=args.end,
            chart_type=args.chart_type,
            output_file=args.output
        )
    except Exception as e:
        print(f"Error: {e}")
        print("Make sure the input file has the correct format with columns: Date, Distance (km), Pace (min/km), Heart Rate (bpm)")

if __name__ == "__main__":
    main()