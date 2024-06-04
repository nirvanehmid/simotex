import streamlit as st
import numpy as np
import pyodbc
import pandas as pd

excel_file_path = "C:/Users/Asus/Desktop/pfatool/end.xlsx"
dataset = pd.read_excel(excel_file_path)

from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier

X = dataset.drop(columns=['absence'])
y = dataset['absence']
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
rf_classifier = RandomForestClassifier()
rf_classifier.fit(X_train, y_train)
conn = pyodbc.connect(
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\Asus\Desktop\pfatool\db.accdb;'
)
def fetch_data(service=None):
    cursor = conn.cursor()
    if service:
        query = f"SELECT matricule, enfant, service, age, anciennete, distanceKM, absence FROM simotex WHERE service = ?"
        cursor.execute(query, (service,))
    else:
        query = "SELECT matricule, enfant, service, age, anciennete, distanceKM, absence FROM simotex"
        cursor.execute(query)
    data = cursor.fetchall()

    return np.array(data)

def add_employee(matricule, enfant, service, age, anciennete, distanceKM, absence=None):
    cursor = conn.cursor()

    # Check if the matricule already exists in the database
    cursor.execute("SELECT COUNT(*) FROM simotex WHERE matricule = ?", (matricule,))
    existing_rows = cursor.fetchone()[0]
    if existing_rows > 0:
        st.error("Matricule déjà existante. Veuillez saisir un matricule différent.")
        return
    else :
        st.success("Employé ajouté avec succès.")

    # Insert the new employee data into the simotex table
    query = """
        INSERT INTO simotex (matricule, enfant, service, age, anciennete, distanceKM, absence)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """
    cursor.execute(query, (matricule, enfant, service, age, anciennete, distanceKM, absence))

    # Commit the transaction
    conn.commit()


def delete_employee(matricule):
    cursor = conn.cursor()

    # Check if the matricule exists in the database
    cursor.execute("SELECT COUNT(*) FROM simotex WHERE matricule = ?", (matricule,))
    existing_rows = cursor.fetchone()[0]

    if existing_rows == 0:
        # Matricule not found
        st.error("Matricule non trouvé.")
        return

    # Delete the employee data from the simotex table
    query = "DELETE FROM simotex WHERE matricule = ?"
    cursor.execute(query, (matricule,))

    # Commit the transaction
    conn.commit()

    # Matricule deleted successfully
    st.success("Matricule supprimé avec succès.")


def add_data_to_excel(enfant=None, service=None, age=None, anciennete=None, distanceKM=None, mois=None, ANNEE=None,
                      absence=None):
    global dataset

    # Mapping encoded values
    enfant_scaled = scaled_enfant.get(enfant, None)
    service_encoded = encoded_service.get(service, None)
    age_scaled = scaled_age.get(age, None)
    anciennete_scaled = scaled_anciennete.get(anciennete, None)
    mois_encoded = encoded_mois.get(mois.lower(), None)  # Lowercase mois for case insensitivity
    ANNEE_scaled = scaled_annee.get(int(ANNEE), None)

    # Check if any of the mapped values are None
    if None in [enfant_scaled, service_encoded, age_scaled, anciennete_scaled, mois_encoded, ANNEE_scaled]:
        return False, "Valeurs non valides."

    # Create a new DataFrame with the new data, following the specified order of columns
    new_data = pd.DataFrame([[enfant_scaled, service_encoded, age_scaled, anciennete_scaled, distanceKM, mois_encoded,
                              ANNEE_scaled, absence]],
                            columns=['enfant', 'service', 'age', 'ancienneté', 'distanceKM', 'mois', 'ANNEE',
                                     'absence'])

    # Concatenate the new data with the existing dataset
    dataset = pd.concat([dataset, new_data], ignore_index=True)

    # Save the updated dataset back to the Excel file
    dataset.to_excel(excel_file_path, index=False)

    return True, "Données ajoutées avec succès."


# Initialisation de l'état de la session si nécessaire
if 'page' not in st.session_state:
    st.session_state.page = 'page1'
encoded_service = {
        'DIRECT PREPARATION': 0.114479,
        'CADRE': 0.019583,
        'DIRECT PACKAGING': 0.131167,
        'WAHIBA DIRECT MONTAGE': 0.098467,
        'IND COLLECTION GAMME TAILLE': 0.107891,
        'STRUCTURE COUPE': 0.052894,
        'INDRECT COUPE': 0.109315,
        'STRUCTURE STOCK ET LOGISTIQUE': 0.023087,
        'STRUCTURE QUALITE': 0.037562,
        'BASMA DIRECT MONTAGE': 0.136804,
        'GENERAUX': 0.038309,
        'DIR COUPE LAIZE': 0.125363,
        'ENCADREMENT COLL G T': 0.043692,
        'MEHERZIA DIRECT MONTAGE': 0.115942,
        'DIRECT QUAL P FINIS': 0.109804,
        'MAGASIN MATIERE PREMIERE': 0.042336,
        'STRUCTURE METHODES': 0.046816,
        'STRUCTURE MECANIQUE ATEL': 0.091795,
        'WAFA DIRECT MONTAGE': 0.100156,
        'MECANIQUE COLLECTION': 0.086379,
        '* STRUCTURE ATELIER': 0.078969,
        'INDIRECT QUALITE PRODUIT FINI': 0.062287,
        'INDIRECT QUALITE MATIERE': 0.097201,
        'DIRECT COUPE PRESSE': 0.135888,
        'CHEDIA DIRECT MONTAGE': 0.100658,
        'EXPEDITION': 0.082088,
        'NAJET DIRECT MONTAGE': 0.106498,
        'FERIDA DIRECT MONTAGE': 0.159344,
        'THOURAYA DIRECT MONTAGE': 0.127470,
        'AZIZA DIRECT MONTAGE': 0.131253,
        '* DISTRIBUTRICE': 0.093091,
        'INDIRECT QUALITE': 0.128010,
        'AMEL DIRECT MONTAGE': 0.094165,
        'STRUCTURE ADMINISTRATION': 0.067311,
        'WAFA SL DIRECT MONTAGE': 0.073634,
        'FORMATION ATELIER': 0.058960,
        'IND EXPEDITION': 0.063814
    }

encoded_mois = {
        'janvier': 0.085277,
        'fevrier': 0.079127,
        'mars': 0.096126,
        'avril': 0.088328,
        'mai': 0.099396,
        'juin': 0.103449,
        'juillet': 0.095669,
        'aout': 0.089828,
        'septembre': 0.127906,
        'octobre': 0.110850,
        'novembre': 0.111442,
        'decembre': 0.102327
    }

scaled_enfant = {
        0: 0.00,
        1: 0.25,
        2: 0.50,
        3: 0.75,
        4: 1.00,
        5: 1.25,
        6: 1.50,
        7: 1.75,
        8: 2.00,
        9: 2.25,
        10: 2.50
    }

scaled_annee = {
    2022: 0.0,
    2023: 1.0,
    2024: 2.0,
    2025: 3.0,
    2026: 4.0,
    2027: 5.0,
    2028: 6.0,
    2029: 7.0,
    2030: 8.0,
    2031: 9.0,
    2032: 10.0,
    2033: 11.0,
    2034: 12.0,
    2035: 13.0,
    2036: 14.0,
    2037: 15.0,
    2038: 16.0,
    2039: 17.0,
    2040: 18.0,
    2041: 19.0,
    2042: 20.0,
    2043: 21.0,
    2044: 22.0,
    2045: 23.0,
    2046: 24.0,
    2047: 25.0,
    2048: 26.0,
    2049: 27.0,
    2050: 28.0,
    2051: 29.0,
    2052: 30.0,
    2053: 31.0,
    2054: 32.0,
    2055: 33.0,
    2056: 34.0,
    2057: 35.0,
    2058: 36.0,
    2059: 37.0,
    2060: 38.0,
    2061: 39.0,
    2062: 40.0,
    2063: 41.0,
    2064: 42.0,
    2065: 43.0,
    2066: 44.0,
    2067: 45.0,
    2068: 46.0,
    2069: 47.0,
    2070: 48.0,
    2071: 49.0,
    2072: 50.0,
    2073: 51.0
}


scaled_anciennete = {
        0: 0.0,
        1: 0.0238,
        2: 0.0476,
        3: 0.0714,
        4: 0.0952,
        5: 0.1190,
        6: 0.1429,
        7: 0.1667,
        8: 0.1905,
        9: 0.2143,
        10: 0.2381,
        11: 0.2619,
        12: 0.2857,
        13: 0.3095,
        14: 0.3333,
        15: 0.3571,
        16: 0.3810,
        17: 0.4048,
        18: 0.4286,
        19: 0.4524,
        20: 0.4762,
        21: 0.5,
        22: 0.5238,
        23: 0.5476,
        24: 0.5714,
        25: 0.5952,
        26: 0.6190,
        27: 0.6429,
        28: 0.6667,
        29: 0.6905,
        30: 0.7143,
        31: 0.7381,
        32: 0.7619,
        33: 0.7857,
        34: 0.8095,
        35: 0.8333,
        36: 0.8571,
        37: 0.8810,
        38: 0.9048,
        39: 0.9286,
        40: 0.9524,
        41: 0.9762,
        42: 1.0
    }
scaled_age = {
        17: 0.0000,
        18: 0.0238,
        19: 0.0476,
        20: 0.0714,
        21: 0.0952,
        22: 0.1190,
        23: 0.1429,
        24: 0.1667,
        25: 0.1905,
        26: 0.2143,
        27: 0.2381,
        28: 0.2619,
        29: 0.2857,
        30: 0.3095,
        31: 0.3333,
        32: 0.3571,
        33: 0.3810,
        34: 0.4048,
        35: 0.4286,
        36: 0.4524,
        37: 0.4762,
        38: 0.5000,
        39: 0.5238,
        40: 0.5476,
        41: 0.5714,
        42: 0.5952,
        43: 0.6190,
        44: 0.6429,
        45: 0.6667,
        46: 0.6905,
        47: 0.7143,
        48: 0.7381,
        49: 0.7619,
        50: 0.7857,
        51: 0.8095,
        52: 0.8333,
        53: 0.8571,
        54: 0.8810,
        55: 0.9048,
        56: 0.9286,
        57: 0.9524,
        58: 0.9762,
        59: 1.0000,
        60: 1.0238,
        61: 1.0476,
        62: 1.0714
    }


# Fonction pour afficher la première page
def page1():
    st.title('absentéisme SIMOTEX')
    st.write("<span style='font-size: small; font-family: Arial;'>", unsafe_allow_html=True)

    # Create a container with three columns for the select boxes
    col1, col2, col3 = st.columns(3)

    # Fetch all data initially to display
    with col1:
        Service = st.selectbox('Service', [
            'DIRECT PREPARATION', 'CADRE', 'DIRECT PACKAGING', 'WAHIBA DIRECT MONTAGE',
            'IND COLLECTION GAMME TAILLE', 'STRUCTURE COUPE', 'INDRECT COUPE',
            'STRUCTURE STOCK ET LOGISTIQUE', 'STRUCTURE QUALITE',
            'BASMA DIRECT MONTAGE', 'GENERAUX', 'DIR COUPE LAIZE',
            'ENCADREMENT COLL G T', 'MEHERZIA DIRECT MONTAGE', 'DIRECT QUAL P FINIS',
            'MAGASIN MATIERE PREMIERE', 'STRUCTURE METHODES',
            'STRUCTURE MECANIQUE ATEL', 'WAFA DIRECT MONTAGE', 'MECANIQUE COLLECTION',
            '* STRUCTURE ATELIER', 'INDIRECT QUALITE PRODUIT FINI',
            'INDIRECT QUALITE MATIERE', 'DIRECT COUPE PRESSE', 'CHEDIA DIRECT MONTAGE',
            'EXPEDITION', 'NAJET DIRECT MONTAGE', 'FERIDA DIRECT MONTAGE',
            'THOURAYA DIRECT MONTAGE', 'AZIZA DIRECT MONTAGE', '* DISTRIBUTRICE',
            'INDIRECT QUALITE', 'AMEL DIRECT MONTAGE', 'STRUCTURE ADMINISTRATION',
            'WAFA SL DIRECT MONTAGE', 'FORMATION ATELIER', 'IND EXPEDITION'], key="service")

    with col2:
        Année = st.selectbox('Année ', [
            '2024', '2025', '2026', '2027', '2028', '2029', '2030', '2031', '2032', '2033', '2034', '2035',
            '2036', '2037', '2038', '2039', '2040', '2041', '2042', '2043', '2044', '2045', '2046', '2047',
            '2048', '2049', '2050', '2051', '2052', '2053', '2054', '2055', '2056', '2057', '2058', '2059',
            '2060', '2061', '2062', '2063', '2064', '2065', '2066', '2067', '2068', '2069', '2070'], key="year")

    with col3:
        Mois = st.selectbox('Mois:', [
            'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Aout', 'Septembre', 'Octobre',
            'Novembre', 'Décembre'], key="month")

    data = fetch_data(Service)
    df = pd.DataFrame(data, columns=['Matricule', 'Enfant', 'Service', 'Age', 'Ancienneté', 'Distance KM', 'Absence'])

    # Display data as pandas DataFrame with increased width and centered
    st.markdown("<h1 style='text-align: center;'></h1>", unsafe_allow_html=True)
    st.dataframe(df, width=800)

    col1, col2, col3 = st.columns([1, 1, 1])

    # Button to add employee (no label)
    with col1:
        show_popup = True

        # Check if the pop-up window should be shown
        if show_popup:
            with st.form("add_delete_employee_form"):
                # Button to add employee
                if st.form_submit_button("Ajouter employé"):
                    st.session_state.page = 'page2_add'

                # Text input for matricule to delete and button to delete employee
                matricule_to_delete = st.text_input("Matricule à supprimer")
                submit_button = st.form_submit_button("Supprimer")

                # Check if the form has been submitted
                if submit_button:
                    # Call the delete function here
                    delete_employee(matricule_to_delete)

    # Button for predicting absence
    with col3:
        # Button for predicting absence
        with col3:
            if st.button("Prédire"):
                # Store selected service, année, and mois in session state
                st.session_state.selected_service = Service
                st.session_state.selected_annee = Année
                st.session_state.selected_mois = Mois

                # Navigate to page 3 for prediction
                st.session_state.page = 'page3_predict'
    from streamlit import form
    st.write(
        "Ce segment permet d'incorporer de nouvelles données réelles dans le jeu de données afin d'améliorer la précision des prédictions.")

    with st.form("my_form"):
        enfant = st.number_input("Nombre d'enfants", min_value=0)
        service = st.selectbox('Service', [
            'DIRECT PREPARATION', 'CADRE', 'DIRECT PACKAGING', 'WAHIBA DIRECT MONTAGE',
            'IND COLLECTION GAMME TAILLE', 'STRUCTURE COUPE', 'INDRECT COUPE',
            'STRUCTURE STOCK ET LOGISTIQUE', 'STRUCTURE QUALITE',
            'BASMA DIRECT MONTAGE', 'GENERAUX', 'DIR COUPE LAIZE',
            'ENCADREMENT COLL G T', 'MEHERZIA DIRECT MONTAGE', 'DIRECT QUAL P FINIS',
            'MAGASIN MATIERE PREMIERE', 'STRUCTURE METHODES',
            'STRUCTURE MECANIQUE ATEL', 'WAFA DIRECT MONTAGE', 'MECANIQUE COLLECTION',
            '* STRUCTURE ATELIER', 'INDIRECT QUALITE PRODUIT FINI',
            'INDIRECT QUALITE MATIERE', 'DIRECT COUPE PRESSE', 'CHEDIA DIRECT MONTAGE',
            'EXPEDITION', 'NAJET DIRECT MONTAGE', 'FERIDA DIRECT MONTAGE',
            'THOURAYA DIRECT MONTAGE', 'AZIZA DIRECT MONTAGE', '* DISTRIBUTRICE',
            'INDIRECT QUALITE', 'AMEL DIRECT MONTAGE', 'STRUCTURE ADMINISTRATION',
            'WAFA SL DIRECT MONTAGE', 'FORMATION ATELIER', 'IND EXPEDITION'])
        age = st.number_input("Age", min_value=0)
        anciennete = st.number_input("Ancienneté", min_value=0)
        distanceKM = st.number_input("Distance (KM)", min_value=0)
        ANNEE = st.selectbox('Année: ', [
            '2024', '2025', '2026', '2027', '2028', '2029', '2030', '2031', '2032', '2033', '2034', '2035',
            '2036', '2037', '2038', '2039', '2040', '2041', '2042', '2043', '2044', '2045', '2046', '2047',
            '2048', '2049', '2050', '2051', '2052', '2053', '2054', '2055', '2056', '2057', '2058', '2059',
            '2060', '2061', '2062', '2063', '2064', '2065', '2066', '2067', '2068', '2069', '2070'])
        mois = st.selectbox('Mois:', [
            'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Aout', 'Septembre', 'Octobre',
            'Novembre', 'Décembre'])
        absence = st.selectbox('Absence', ['Low', 'High'])

        submitted = st.form_submit_button("Ajouter")
        if submitted:
            success, message = add_data_to_excel(enfant, service, age, anciennete,distanceKM, mois, ANNEE, absence)
            if success:
                st.success(message)
            else:
                st.error(message)


def page3_predict():
    # Fetch data for the selected service
    selected_service = st.session_state.selected_service
    data = fetch_data(selected_service)

    # Create a DataFrame
    df = pd.DataFrame(data, columns=['Matricule', 'Enfant', 'Service', 'Age', 'Ancienneté', 'Distance KM', 'Absence'])

    # Add columns for année, mois, and absence
    df['Année'] = st.session_state.selected_annee
    df['Mois'] = st.session_state.selected_mois
    cols = df.columns.tolist()
    cols.remove('Absence')
    cols.extend(['Absence'])
    df = df[cols]
    # Display the DataFrame with the predicted absence column


    # Encode categorical features
    df['service'] = df['Service'].map(encoded_service)
    df['mois'] = df['Mois'].map(encoded_mois)

    # Scale numerical features
    df['enfant'] = df['Enfant'].map(lambda x: scaled_enfant[x])
    df['ANNEE'] = df['Année'].map(scaled_annee)
    df['ancienneté'] = df['Ancienneté'].map(scaled_anciennete)
    df['age']=df['Age'].map(scaled_age)
    df['distanceKM'] = df['Distance KM']
    # Select the features for prediction
    X_pred = df[['service','distanceKM','mois','age','ancienneté','enfant','ANNEE']]

    # Predict absence using the trained model
    predicted_absence = rf_classifier.predict(X_pred)

    # Assign the predicted values to the 'Absence' column
    df['Absence'] = predicted_absence
    original_columns = ['Matricule', 'Enfant', 'Service', 'Age', 'Ancienneté', 'Distance KM', 'Absence']
    st.title(f"Prédiction pour le service {selected_service}, pour le mois de {st.session_state.selected_mois} de l'année {st.session_state.selected_annee}:")
    st.dataframe(df[original_columns])
    import plotly.express as px
    from sklearn.metrics import accuracy_score, precision_score
    X_test = dataset.drop(columns=['absence'])
    y_test = dataset['absence']

    # Predictions on the test set
    y_pred = rf_classifier.predict(X_test)

    # Evaluation metrics
    accuracy = accuracy_score(y_test, y_pred)
    precision = precision_score(y_test, y_pred, average=None)
    # Data for pie chart
    low_count = (df['Absence'] == 'Low').sum()
    high_count = (df['Absence'] == 'High').sum()

    # Create a DataFrame for the absence counts
    absence_counts = pd.DataFrame({'Absence': ['Low', 'High'], 'Count': [low_count, high_count]})

    # Create the pie chart with specified colors
    fig = px.pie(absence_counts, names='Absence', values='Count', title='Répartition des absences',
                 color='Absence', color_discrete_map={'High': 'red', 'Low': 'skyblue'})

    # Update layout to adjust width and font size
    fig.update_layout(
        autosize=False,
        width=400,  # Adjust width as needed
        height=400,  # Adjust height as needed
        font=dict(
            size=20  # Adjust font size as needed
        )
    )

    # Display the plot and the evaluation metrics horizontally
    col1, col2, col3, col4 = st.columns([2, 2,2,2])
    with col1:
        st.plotly_chart(fig)
    # Button to return to page 1
    if st.button("Retourner à la page principale"):
        st.session_state.page = 'page1'
# Fonction pour afficher la deuxième page pour ajouter un employé
def page2_add():
    st.title('Ajouter un employé')
    st.write('Remplissez les informations pour ajouter un employé à la base de données.')

    # Champ de saisie pour les informations de l'employé
    matricule = st.text_input("Matricule")

    # Validate numerical inputs
    enfant = st.number_input("Nombre d'enfants", min_value=0)
    age = st.number_input("Age", min_value=0)
    anciennete = st.number_input("Ancienneté", min_value=0)
    distanceKM = st.number_input("Distance (KM)", min_value=0)

    service = st.selectbox('Service', [
        'DIRECT PREPARATION', 'CADRE', 'DIRECT PACKAGING', 'WAHIBA DIRECT MONTAGE',
        'IND COLLECTION GAMME TAILLE', 'STRUCTURE COUPE', 'INDRECT COUPE',
        'STRUCTURE STOCK ET LOGISTIQUE', 'STRUCTURE QUALITE',
        'BASMA DIRECT MONTAGE', 'GENERAUX', 'DIR COUPE LAIZE',
        'ENCADREMENT COLL G T', 'MEHERZIA DIRECT MONTAGE', 'DIRECT QUAL P FINIS',
        'MAGASIN MATIERE PREMIERE', 'STRUCTURE METHODES',
        'STRUCTURE MECANIQUE ATEL', 'WAFA DIRECT MONTAGE', 'MECANIQUE COLLECTION',
        '* STRUCTURE ATELIER', 'INDIRECT QUALITE PRODUIT FINI',
        'INDIRECT QUALITE MATIERE', 'DIRECT COUPE PRESSE', 'CHEDIA DIRECT MONTAGE',
        'EXPEDITION', 'NAJET DIRECT MONTAGE', 'FERIDA DIRECT MONTAGE',
        'THOURAYA DIRECT MONTAGE', 'AZIZA DIRECT MONTAGE', '* DISTRIBUTRICE',
        'INDIRECT QUALITE', 'AMEL DIRECT MONTAGE', 'STRUCTURE ADMINISTRATION',
        'WAFA SL DIRECT MONTAGE', 'FORMATION ATELIER', 'IND EXPEDITION'], key="service")

    col1, col2 = st.columns(2)
    with col1:
        with st.form("add_employee_form"):
            if st.form_submit_button("Ajouter employé"):
                # Check if any input field is empty
                if matricule == "":
                    st.error("Veuillez remplir tous les champs.")
                else:
                    # Add employee to database
                    add_employee(matricule, enfant, service, age, anciennete, distanceKM)

        with col2:
            if st.button("Retourner à la page principale"):
                # Return to page 1 without adding an employee
                st.session_state.page = 'page1'



# Affichage de la page en fonction de l'état act
if st.session_state.page == 'page1':
    page1()
elif st.session_state.page == 'page2_add':
    page2_add()
elif st.session_state.page == 'page3_predict':
    page3_predict()