# Comments:
'''
- None.
'''

# Libraries:
from random import random
from os import listdir
from os.path import exists, join
from pandas import ExcelFile, ExcelWriter, read_excel, DataFrame, concat, pivot_table
from matplotlib.pyplot import subplots, show
from scipy.odr import ODR, Model, RealData 
from scipy.optimize import curve_fit
# Classes:

# Functions:
## Welcome the user to the program and explain it:
def intro():
    print('-------------------------')
    print('Welcome to "Solution Analyser"!')
    print('-------------------------')
    print()
    print('Here you will be able to:')
    print('1. Get the specific rotation of carbohydrates;')
    print('2. Get the concentration profile of inhomogeneous water-carbohydrate solutions;')
    print('3. Convert the concentration to refractive index profile.')
    print()
    print("Let's start!")
    print()
## Ask if the user wants to save the graphics and tables:
def ask_save():
    global save
    # Aks if the user wants to save the graphics and tables:
    while True:
        question = input('Do you want to save plots and tables?: [y/n]')
        try:
            if question not in ['y', 'n']: # Invalid answer.
                raise ValueError('Invalid answer!')
            elif question == 'y': # Positive answer.
                save = True
            else: # Negative answer.
                save = False 
        except ValueError as ve: # It occured an error.
            print(ve)
        else: # It ran smoothly.
            break
## Find the path to the data: 
def get_folder():
    global folder
    # Get the folder path from the txt file path manager:
    while True:
        with open("Path Manager.txt", mode='rt') as f:
            folder = f.read()
            f.close()
        try: 
            if not exists(folder):
                raise FileExistsError('There is no folder with such path!')
            else:
                break
        except FileExistsError as fee:
            print(fee)
            print('Go to the "Path Manager.txt" and change its content!')
            input('Type anything when you are done: ')
## Get the aquarium length:
def ask_L():
    # Aks if the user wants to save the graphics and tables:
    global L
    while True:
        question = input('Lenght of the aquarium [dm]: [6]')
        try:
            if not question.replace('.','').isdigit():
                raise ValueError('Not a valid number!')
            L = float(question)
            if L == 0:
                raise ValueError('Null value!')
        except ValueError as ve: 
            print(ve)
        else: 
            break
## Get the general uncertainties:
def get_uncertainties():
    # Length:
    global uL; global uh
    L_scl, h_scl = 1e-2, 1e-2 # dm.
    uL, uh = L_scl / (2 * 6**0.5), h_scl / (2 * 6**0.5)
    # Volume:
    global uV
    V_scl = 100 # ml.
    uV = V_scl / (2 * 6**0.5)
    # Mass:
    global um
    m_scl = 0.1 # g.
    um = m_scl / (2 * 6**0.5)
    # Angle:
    global udTheta
    Theta_scl = 2 # °.
    uTheta = Theta_scl / (2 * 6**0.5)
    udTheta = 2**0.5 * uTheta
## Calculate the specific rotation for carbohydrates:
def get_specific_rotation():
    alpha_list = list()
    path = join(folder, 'Data - Calibration.xlsx')
    sheets = ExcelFile(path).sheet_names
    # Figure:
    fig, ax = subplots(nrows=1, ncols=1, layout='constrained', figsize=(6,4))
    ax.set_xlabel('Concentration [g/ml]', loc='center', fontsize=14)
    ax.set_ylabel('Angle displacement [°]', loc='center', fontsize=14)
    ax.tick_params(axis='both', which='major', labelsize=12)
    ## For each sheet:
    df_cal = DataFrame()
    for i, sheet in enumerate(sheets):
        # Reading data:
        df = read_excel(path, sheet_name=sheet)
        C, dTheta = df.loc[:, 'C [g/ml]'].values, df.loc[:, 'dTheta [°]'].values
        # Uncertainties:
        ## Concentration:
        V = 400 # mL.
        uC = (um**2 + (C * uV)**2)**0.5 / V
        # Regression:
        def fit_function(p, x):
            a = p
            return a * x
        model = Model(fit_function)
        data = RealData(C, dTheta, sx=uC, sy=udTheta)
        odr = ODR(data, model, beta0=[0.1])
        result = odr.run()
        a, ua = result.beta[0], result.sd_beta[0]
        dTheta_pred = a * C
        alpha = a / L
        ualpha = (ua**2 + (alpha * uL)**2)**0.5 / L
        # Data frame:
        df_cal.loc[i, 'Carbohydrate'] = sheet.capitalize()
        df_cal.loc[i, 'alpha [ml/(g.dm)]'], df_cal.loc[i, 'ualpha [ml/(g.dm)]'] = alpha, ualpha
        # Plot:
        color = (random(), random(), random())
        ax.errorbar(C, dTheta, xerr=uC, yerr=udTheta, marker='.', ms=20, mec='black', mfc=color, linestyle='none', lw=2, ecolor='black', capsize=5, label=f'{sheet.capitalize()}')
        ax.plot(C, dTheta_pred, c=color, lw=2, linestyle='solid')
    ax.legend(loc='lower left', fontsize=12)
    show()
    if save:
        fig.savefig(join(folder, f'Graphic - Calibration.png'))
        with ExcelWriter(join(folder, f'Table - Calibration.xlsx'), engine='openpyxl') as writer:
            df_cal.to_excel(writer, index=False)
    return df_cal
## Concentration analysis for inhomegenous water-carbohydrate solutions:
def handbook_calibration():
    df_han = read_excel(join(folder, f'Data - Handbook.xlsx'))
    C, n = df_han.loc[:, 'C [g/ml]'].values, df_han.loc[:, 'n'].values 
    # Regression:
    def fit_function(x, a, b, c):
        return a * x**2 + b * x + c
    params, covariance = curve_fit(fit_function, C, n)
    a, b, c = params
    ua, ub, uc = covariance[0, 0], covariance[1, 1], covariance[2, 2]
    df_params = DataFrame(data = {'a [ml^2/g^2]': [a], 'ua [ml^2/g^2]': [ua], 'b [ml/g]': [b], 'ub [ml/g]': [ub], 'c': [c], 'uc': [uc]})
    n_pred = fit_function(C, a, b, c)
    # Figure:
    fig, ax = subplots(nrows=1, ncols=1, layout='constrained', figsize=(6,4))
    ax.set_xlabel('Concentration [g/ml]', loc='center', fontsize=14)
    ax.set_ylabel('Refractive index', loc='center', fontsize=14)
    ax.tick_params(axis='both', which='major', labelsize=12)
    ax.scatter(C, n, s=80, c='black', label='Data')
    ax.plot(C, n_pred, c='red', lw=2, label='Fit')
    ax.legend(loc='best', fontsize=12)
    show()
    if save:
        fig.savefig(join(folder, f'Graphic - Handbook.png'))
        with ExcelWriter(join(folder, f'Table - Handbook.xlsx'), engine='openpyxl') as writer:
            df_params.to_excel(writer, index=False)
    # Create the model function with the numerical parameters:
    f, uf = lambda x: a * x**2 + b * x + c, lambda x, ux: abs(2 * a * x + b) * ux
    return f, uf
## Convert concentration data into refractive index:
def analysis(df_han, f, uf):
    path_list = [join(folder, file) for file in listdir(folder) if 'Varying Height' in file]
    for path in path_list:
        # Get the specific rotation coefficient:
        carbohydrate = path.split('-')[-1].replace('.xlsx','').strip()
        alpha, ualpha = df_han.loc[df_han['Carbohydrate'] == carbohydrate, 'alpha [ml/(g.dm)]'].values[0], df_han.loc[df_han['Carbohydrate'] == carbohydrate, 'ualpha [ml/(g.dm)]'].values[0]
        # Figure:
        ## Concentration:
        figC, axC = subplots(nrows=1, ncols=1, layout='constrained', figsize=(6,4))
        axC.set_xlabel('Height [cm]', loc='center', fontsize=14)
        axC.set_ylabel('Concentration [g/ml]', loc='center', fontsize=14)
        axC.tick_params(axis='both', which='major', labelsize=12)
        ## Refractive index:
        fign, axn = subplots(nrows=1, ncols=1, layout='constrained', figsize=(6,4))
        axn.set_xlabel('Height [cm]', loc='center', fontsize=14)
        axn.set_ylabel('Refractive index', loc='center', fontsize=14)
        axn.tick_params(axis='both', which='major', labelsize=12)
        ## Initialization of data frames:
        df_paramsC, df_paramsn, df_time = DataFrame(), DataFrame(), DataFrame(columns = ['Day', 'h [cm]', 'C [g/ml]'])
        ## Read the sheet data:
        sheets = ExcelFile(path).sheet_names
        for (i, sheet) in enumerate(sheets):
            # Reading data:
            day = int(sheet.strip().replace('Day',''))
            df = read_excel(path, sheet_name=sheet)
            h, dTheta = df.loc[:, 'h [cm]'].values, df.loc[:, 'dTheta [°]'].values
            C = dTheta / (alpha * L)
            uC = (udTheta**2 + dTheta**2 * ((ualpha / alpha)**2 + (uL / L)**2))**0.5 / (alpha * L)
            n, un = f(C), uf(C, uC)
            for (h_, C_) in zip(h, C):
                new_df = DataFrame(data={'Day': [day], 'h [cm]': [h_], 'C [g/ml]': [C_]})
                df_time = concat([df_time, new_df], ignore_index=True)
            # Regression:
            def fit_function(p, x):
                a, b, c, d = p
                return a + (b - a) / (1 + (x / c)**d)
            ## Concentration:
            model = Model(fit_function)
            data = RealData(h, C, sx=uh, sy=uC)
            odr = ODR(data, model, beta0=[0.6, 0.1, 4., 6.])
            result = odr.run()
            (a, b, c, d), (ua, ub, uc, ud) = result.beta, result.sd_beta
            ### Storing the regression coefficients:
            df_paramsC.loc[i, 'Day'] = day
            df_paramsC.loc[i, 'a [g/ml]'], df_paramsC.loc[i, 'ua [g/ml]'] = a, ua
            df_paramsC.loc[i, 'b [g/ml]'], df_paramsC.loc[i, 'ub [g/ml]'] = b, ub
            df_paramsC.loc[i, 'c [cm]'], df_paramsC.loc[i, 'uc [cm]'] = c, uc
            df_paramsC.loc[i, 'd'], df_paramsC.loc[i, 'ud'] = d, ud
            C_pred = fit_function((a, b, c, d), h)
            ## Refractive index:
            model = Model(fit_function)
            data = RealData(h, n, sx=uh, sy=un)
            odr = ODR(data, model, beta0=[0.6, 0.1, 4., 6.])
            result = odr.run()
            (a, b, c, d), (ua, ub, uc, ud) = result.beta, result.sd_beta
            n_pred = fit_function((a, b, c, d), h)
            ### Storing the regression coefficients:
            df_paramsn.loc[i, 'Day'] = sheet.strip().replace('Day','')
            df_paramsn.loc[i, 'a'], df_paramsn.loc[i, 'ua'] = a, ua
            df_paramsn.loc[i, 'b'], df_paramsn.loc[i, 'ub'] = b, ub
            df_paramsn.loc[i, 'c [cm]'], df_paramsn.loc[i, 'uc [cm]'] = c, uc
            df_paramsn.loc[i, 'd'], df_paramsn.loc[i, 'ud'] = d, ud
            # Plot:
            color = (random(), random(), random())
            ## Concentration:
            axC.errorbar(h, C, xerr=uh, yerr=uC, marker='.', ms=20, mec='black', mfc=color, linestyle='none', lw=2, ecolor='black', capsize=5, label=sheet)
            axC.plot(h, C_pred, c=color, lw=2, linestyle='solid')
            ## Refractive index:
            axn.errorbar(h, n, xerr=uh, yerr=un, marker='.', ms=20, mec='black', mfc=color, linestyle='none', lw=2, ecolor='black', capsize=5, label=sheet)
            axn.plot(h, n_pred, c=color, lw=2, linestyle='solid')
        axC.legend(loc='best', fontsize=12), axn.legend(loc='best', fontsize=12)
        show()
        # Statistical analysis of the regression coefficients:
        df_statsC = df_paramsC.describe().loc[['mean', 'std'], ['a [g/ml]', 'b [g/ml]', 'c [cm]', 'd']]
        df_statsn = df_paramsn.describe().loc[['mean', 'std'], ['a', 'b', 'c [cm]', 'd']]
        # Time evolution:
        df_time = pivot_table(df_time, values='C [g/ml]', index='h [cm]', columns='Day')
        figt, axt = subplots(nrows=1, ncols=1, layout='constrained', figsize=(6,4))
        axt.set_xlabel('Height [cm]', loc='center', fontsize=14)
        axt.set_ylabel('Concentration [g/ml]', loc='center', fontsize=14)
        axt.tick_params(axis='both', which='major', labelsize=12)
        h_values = df_time.index
        N = len(h_values)
        ind = (N - 1)// 2
        for h in h_values:
            if (h == h_values[0]) or (h == h_values[-1]) or (h == h_values[ind]): 
                # Reading the data:
                dt = df_time.columns.values
                C = df_time.loc[h, :].values
                # Regression:
                def fit_function(x, a, b):
                    return a * x + b
                params, covariance = curve_fit(fit_function, dt, C)
                a, b = params
                ua, ub = covariance[0, 0], covariance[1, 1]
                C_pred = fit_function(dt, a, b)
                # Plot:
                color = (random(), random(), random())
                axt.scatter(dt, C, s=80, c=color, label=h)
                axt.plot(dt, C_pred, c=color, lw=2, linestyle='solid')
        axt.legend(title='h [cm]', loc='best', fontsize=12, title_fontsize=12)
        show()
        # Save the plots and tables:
        if save:
            # Concetration:
            figC.savefig(join(folder, f'Graphic - Concentration Profile - {carbohydrate}.png'))
            with ExcelWriter(join(folder, f'Table - Concentration Profile - {carbohydrate}.xlsx'), engine='openpyxl') as writer:
                df_paramsC.to_excel(writer, sheet_name='Params', index=False)
                df_statsC.to_excel(writer, sheet_name='Stats', index=False)
            # Refractive Index:
            fign.savefig(join(folder, f'Graphic - Refractive Index Profile - {carbohydrate}.png'))
            with ExcelWriter(join(folder, f'Table - Refractive Index Profile - {carbohydrate}.xlsx'), engine='openpyxl') as writer:
                df_paramsn.to_excel(writer, sheet_name='Params', index=False)
                df_statsn.to_excel(writer, sheet_name='Stats', index=False)
            # Time evlolution:
            figt.savefig(join(folder, f'Graphic - Concentration Time Evolution.png'))
## Ask if the user wants to stop the program: 
def ask_stop():
    while True:
        # Aks about the stop:
        question = input('Do you want to stop the program? [y/n]')
        try:
            if question not in ['y', 'n']: # Invalid answer.
                raise ValueError('Invalid answer!')
            elif question == 'y':  # Will stop.
                stop = True
            else: # Won't stop.
                stop = False
        except ValueError as ve: # It occured an error.
            print(ve)
        else: # It ran smoothly.
            break
    return stop        
## Main control function:
def main():
    # Welcome the user to the program and explain it:
    intro()
    # Ask if the user wants to save graphics and tables:
    save = ask_save()
    # Get the folder with the data:
    get_folder()
    # Get the aquarium info:
    ask_L()
    # Get the general uncertainties:
    get_uncertainties()
    # Main loop:
    stop = False
    while not stop:
        # Get the specific rotation of carbohydrates:
        df_han = get_specific_rotation()
        # Conventration analysis:
        f, uf = handbook_calibration()
        # Main analysis:
        analysis(df_han, f, uf)
        # Ask if the user wants to interrupt the program:
        stop = ask_stop()

# Run the program:
main()