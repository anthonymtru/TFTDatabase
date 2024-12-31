import pandas as pd

df = pd.read_excel('tft_data.xlsx')
games_count = df['Placement'].count()

def winrate_stat():
    wins_count = df['Win'].sum()
    percent = (wins_count * 100.0)/games_count
    return 'Your win rate is ' + str(percent) + '%.'

def top4_rate_stat():
    top4_count = df['Top 4'].sum()
    percent = (top4_count * 100.0)/games_count
    return 'Your top 4 rate is ' + str(percent) + '%.'

def augment_stats():
    user_input = input('What augment would you like to check? ').lower()
    # Filter the DataFrame where the augment appears in any of the three augment columns
    filtered_df = df[
        (df['Augment 1'].str.lower() == user_input) |
        (df['Augment 2'].str.lower() == user_input) |
        (df['Augment 3'].str.lower() == user_input)
    ]
    filtered_games = filtered_df['Win'].count()
    
    user_input = input('What would you like to check? (Top4Rate/Winrate/AvgPlacement) ').lower()
    if user_input == 'top4rate':
        top4_count = filtered_df['Top 4'].sum()
        percent = (top4_count * 100.0)/filtered_games
        return 'Your top 4 rate with ' + user_input + ' is ' + str(percent) + '%.'
    elif user_input == 'winrate':
        win_count = filtered_df['Win'].sum()
        percent = (win_count * 100.0)/filtered_games
        return 'Your win rate with ' + user_input + ' is ' + str(percent) + '%.'
    elif user_input == 'avgplacement':
        placements = filtered_df['Placement'].sum()
        percent = (placements * 1.0)/filtered_games
        return 'Your average placement for ' + user_input + ' is ' + str(percent) + '.'
    else:
        print('There was an error in your input, please redo the process.')

# Run the main function
if __name__ == '__main__':
    while True:
        user_input = input('What would you like to check? (check list attached to readme) ').lower()
        if 'winrate' in user_input:
            print(winrate_stat())
            break
        elif 'top4rate' in user_input:
            print(top4_rate_stat())
            break
        elif 'augmentstats' in user_input:
            print(augment_stats())
            break
        else:
            print('Invalid option. Please refer back to the readme file.')