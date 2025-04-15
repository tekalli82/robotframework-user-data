***Settings***
Library           ./UserData.py
Library           OperatingSystem

***Variables***
${API_URL}        https://jsonplaceholder.typicode.com/users
${OUTPUT_DIR}     ./

***Test Cases***
Hae Ja Tallenna Käyttäjät Exceliin
    ${excel_filename}=    Create Excel Filename
    ${json_file_exists}=    Run Keyword And Return Status    OperatingSystem.File Should Exist    users.json
    IF    ${json_file_exists}
        Log    Käyttäjädata löytyy tiedostosta, ei tehdä rajapintakutsua.
        ${users}=    Load Users Data From File
    ELSE
        Log    Käyttäjädataa ei löydy tiedostosta, haetaan rajapinnasta.
        ${users}=    Fetch Users From Api    ${API_URL}
        Save Users To File    ${users}
    END
    ${processed_users}=    Process User Data    ${users}
    ${sorted_users}=    Sort Users    ${processed_users}
    ${output_path}=    OperatingSystem.Join Path    ${OUTPUT_DIR}    ${excel_filename}
    Save To Excel    ${sorted_users}    ${output_path}
    Log    Excel-tiedostonimi luotu: ${excel_filename} # Testi
	
***Keywords***
Create Excel Filename
    ${filename}=    Evaluate    UserData.create_excel_filename()
    RETURN    ${filename}

Load Users Data From File
    ${users}=    Load Users From File
    RETURN    ${users}