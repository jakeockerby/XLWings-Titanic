import xlwings as xw
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import datetime
from time import sleep
from matplotlib import cm
from tpot import TPOTClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import confusion_matrix


def bar_labels(ax, space=0.8, fontsize=12):
    for p in ax.patches:
        x = p.get_x() + p.get_width() / 2 # Plotting at centre of bar
        y = p.get_y() + p.get_height() + float(space) # Plotting halo just above bar
        
        # Storing bar height value as an integer
        value = int(p.get_height())
        
        # Align at center with the fontsize passed
        ax.text(x, y, value, ha="center", fontsize=fontsize)


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    # Storing sheet values as variables
    age1 = int(sheet["B2"].value)
    age2 = int(sheet["B3"].value)
    sex = sheet['B5'].value
    class_ = sheet["B7"].value
    survived = sheet["B9"].value
    t_length = sheet["B11"].value
    sort = sheet["B12"].value
    asc = sheet["B13"].value
    graphs = sheet["B15"].value
    title = sheet["B17"].value
    title_font = sheet["B18"].value
    tfontsize = sheet["B19"].value
    halos = sheet["B21"].value
    xaxis = str(sheet["B23"].value)
    yaxis = str(sheet["B24"].value)
    hue = sheet["B25"].value
    export = sheet["B27"].value
    loc = sheet["B29"].value
    
    # Converting ascending variable into boolean values for later
    if asc == 'Yes':
        asc = True
    else:
        asc = False

    
    # Loading the titanic dataset
    url = r'https://raw.githubusercontent.com/gmonce/scikit-learn-book/master/data/titanic.csv'
    
    # Only getting columns we want
    titanic = pd.read_csv(url, usecols=["pclass","survived","name","age", 
                                        "sex"])
    
    return titanic


    # Replacing 'male' and 'female' to match the values in the excel sheet
    titanic['sex'] = titanic['sex'].replace('male', 'Male')
    titanic['sex'] = titanic['sex'].replace('female', 'Female')
    
    # Converting survived values 'Yes' and 'No' to 1 and 0 respectively
    titanic['survived'] = titanic['survived'].replace(0, 'No')
    titanic['survived'] = titanic['survived'].replace(1, 'Yes')
    
    # Filling NaN values with the mean age
    titanic['age'] = titanic['age'].fillna(int(round(titanic['age'].mean())))
    
    # Filter by ages the user has selected
    titanic = titanic.loc[(titanic['age'] >= age1) & (titanic['age'] <= age2)]
    
    # Filter depending on the characteristics the user has selected
    if sex != 'Both':
        titanic = titanic.loc[titanic['sex'] == sex]
    if class_ != 'All':
        titanic = titanic.loc[titanic['pclass'] == class_]
    if survived != 'Both':
        titanic = titanic.loc[titanic['survived'] == survived]
    
    # Sorting values by the filters the user has selected for the sort and asc
    # variables
    titanic = titanic.sort_values(by=str(sort), ascending=asc)
    
    pivot = titanic
    pivot['survived'] = pivot['survived'].replace('No', 0)
    pivot['survived'] = pivot['survived'].replace('Yes', 1)
    pivot = pd.pivot_table(pivot, values=yaxis, index=[xaxis], columns=[hue],
                           aggfunc='sum')

    # Graphs
    if graphs == 'Bar Graph':
        graph = pivot.plot(kind='bar', figsize=(14, 6), 
                           ylabel='Number Survived', colormap='Set2')
        
        # Despine graph
        sns.despine(left=True)
        
        # Plot title with the title, font and fontsize user has entered
        tfont = {'fontname':'{}'.format(title_font)}
        plt.title('{}'.format(title),**tfont, fontsize=tfontsize, y=1.1)
        
        # Get current axes and use the function defined above to add halos
        ax = plt.gca()
        if halos == 'On':
            bar_labels(ax)
        
        # Getting current time for filenames
        current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
        
        # Save graphs
        graph.figure.savefig('{0}titanic_bar_{1}.png'.format(export,
                      current_time),
                      transparent=True)
    
    
    if graphs == 'Pie Chart':
        # Transposing the pivot table so that the pie chart works properly
        pivot_t = pivot.T
        
        # Adding a count to add on the end of filenames if multiple charts are
        # produced
        c = 1

        # For every index in the original pivot table
        for idx in pivot.index:
            plt.figure(figsize=(20,6))
            
            # Setting the colour scheme
            colors = cm.Set2(np.arange(len(pivot_t))/len(pivot_t))
            
            # Fonts for title and labels
            tfont = {'fontname':'{}'.format(title_font)}
            
            # Set user inputted title with correct formating
            if len(pivot.index) == 1:
                # If there's only 1 category, keep original title
                plt.title('{}'.format(title),**tfont, fontsize=tfontsize,
                          y=1.1)
            else:
                # Else add the index on the end so it is clear which category
                # the pie belongs to
                plt.title('{0} - {1}'.format(title, idx),**tfont,
                          fontsize=tfontsize, y=1.1)
            
            # Creating pie chart with percentages in the middle
            pie, q, w = plt.pie(pivot_t[idx], autopct='%1.0f%%',
                    pctdistance=0.75, labeldistance=1.35, counterclock=False,
                    colors=colors, startangle=90)
            
            # Plotting a legend            
            plt.legend(loc='center', bbox_to_anchor=(1.2, 0.5),
                       labels=pivot_t.index, frameon=False)
            
             
            p = plt.gcf()
            
            # Placing a circle in the middle of the pie to create a donut
            plt.setp(pie, width=0.5)
            
            # Getting current time for filenames
            current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
            
            # Save charts
            p.savefig('{0}titanic_pie_{1}_{2}.png'.format(export,
                      current_time, c),
                      transparent=True)
            
            # Adding 1 to the count
            c += 1
    
    # If table length is restricted by the user, get first t_length rows of the
    # dataframe
    if t_length != 'display full':
        t_length = int(t_length)
        titanic = titanic.head(t_length)
        
    # Print the final table to the excel sheet
    sheet["{}".format(loc)].value = titanic




def predictions():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    
    loc = sheet["B29"].value
    export = sheet["B27"].value
    
    titanic = main()
    titanic = titanic.drop('name', axis=1)
    titanic['sex'] = titanic['sex'].replace('male', 0)
    titanic['sex'] = titanic['sex'].replace('female', 1)
    
    titanic['pclass'] = titanic['pclass'].replace('1st', 0)
    titanic['pclass'] = titanic['pclass'].replace('2nd', 1)
    titanic['pclass'] = titanic['pclass'].replace('3rd', 2)
    
    # Filling NaN values with the mean age
    titanic['age'] = titanic['age'].fillna(int(round(titanic['age'].mean())))
    

    training = ['pclass', 'age', 'sex']
    
    X = titanic[training].values
    y = titanic['survived'].values
    
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3,
                                       random_state=42)
    
    tpot = TPOTClassifier(generations=10, verbosity=2,
                          random_state=42, n_jobs=-1)
    tpot.fit(X_train, y_train)
    
    sheet["{}".format(loc)].value = 'Training Accuracy:'
    sheet["{}".format(loc)].offset(0, 1).value = tpot.score(X_train, y_train)
    sheet["{}".format(loc)].offset(2, 0).value = 'Test Accuracy:'
    sheet["{}".format(loc)].offset(2, 1).value = tpot.score(X_test, y_test)

    y_pred = tpot.predict(X)

    mat = confusion_matrix(y, y_pred)
    
    mat_df = pd.DataFrame(mat)
    fig, ax = plt.subplots()
    plt.imshow(mat_df, cmap="winter", interpolation='nearest')
    plt.axis('off')
    for i in range(2):
        for j in range(2):
            text = ax.text(j, i, mat[i, j],
                            ha="center", va="center", color="w")
    # Getting current time for filenames
    current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
    plt.savefig('{0}confusion_matrix_{1}.png'.format(export, current_time))
    
# Function to clear the tables
def clear():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet.range('A30:B100000').clear()
    sheet.range('C1:ZZ100000').clear()
