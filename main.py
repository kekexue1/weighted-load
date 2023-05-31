# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


from rostering_flask.models import User, Pool, Institution, Qualification, PhysicianQualification, Station, PhysicianStation, DutyType, DutyBlockType, PlanningPeriod, DutyBlockTypeAssignment, PreferencesForDaysOptionsConfiguration, GeneralPreferencesConfiguration, DefaultValuePreferenceOptions, TeamType, TeamQualificationType, TeamBlockType, TeamBlockTypeAssignment, PhysicianPool, DutyQualificationType, PoolDutyTypeAssignment
from flask_sqlalchemy import SQLAlchemy
from rostering_flask import db, bcrypt
from rostering_flask import create_app
import datetime
import os
import pandas as pd


path = os.path.dirname(__file__) + "/PlanningData.xlsx"
table_users = pd.read_excel(path, sheet_name='Nutzer')
table_pools = pd.read_excel(path, sheet_name='Pools')
table_teams = pd.read_excel(path, sheet_name='Stationen')
table_duties = pd.read_excel(path, sheet_name='Dienste')
table_team_duties = pd.read_excel(path, sheet_name='Stationsdienste')

app = create_app()

db.drop_all()

db.create_all()

institution_klinikum_straubing = Institution(name="Klinikum Straubing")

qualification_list_string = list()
for user in table_users.values:
    if user.tolist()[list(table_users).index("Qualifikationen")] == user.tolist()[list(table_users).index("Qualifikationen")]:
        for quali in user.tolist()[list(table_users).index("Qualifikationen")].replace(" ", "").split(','):
            qualification_list_string.append(quali)

qualification_list_string = list(dict.fromkeys(qualification_list_string))
for quali in qualification_list_string:
    new_quali = Qualification(name=quali, institution_id=1)
    db.session.add(new_quali)


for user in table_users.values:
    email = user.tolist()[list(table_users).index("E-Mail")]
    first_name = user.tolist()[list(table_users).index("Vorname")]
    last_name = user.tolist()[list(table_users).index("Nachname")]
    # role = user.tolist()[list(table_users).index("Planungsberechtigung")]
    role = "admin"
    planned_manually = user.tolist()[list(table_users).index("Wird manuell geplant")]
    if planned_manually == "ja":
        planned_manually = True
    else:
        planned_manually = False

    # beschaeftigungsumfang = user.tolist()[list(table_users).index("Beschäftigungsumfang in %")]
    beschaeftigungsumfang = 100
    is_being_planned = user.tolist()[list(table_users).index("Nur für die Verwaltung")]
    if is_being_planned == "ja":
        is_being_planned = False
    else: is_being_planned = True

    new_user = User(email=email, password=bcrypt.generate_password_hash("init1234").decode("utf-8"), first_name=first_name, last_name=last_name, role=role, beschaeftigungsumfang=beschaeftigungsumfang, planned_manually=planned_manually, is_being_planned=is_being_planned, institution_id=1)
    db.session.add(new_user)

    if user.tolist()[list(table_users).index("Qualifikationen")] == user.tolist()[list(table_users).index("Qualifikationen")]:
        qualis_strings = user.tolist()[list(table_users).index("Qualifikationen")].replace(" ", "").split(',')
        for quali_string in qualis_strings:
            new_physician_quali = PhysicianQualification(institution_id = 1, physician_id = [x for x in User.query.all() if x == new_user][0].id, qualification_id = [x for x in Qualification.query.all() if x.name == quali_string][0].id)
            db.session.add(new_physician_quali)


users = User.query.filter_by(is_being_planned=1).all()
users.sort(key=lambda x: x.last_name, reverse=False)
list_of_physicians = users

for team in table_teams.values:
    name = team.tolist()[list(table_teams).index("Stationsname")]
    new_team = Station(name=name, institution_id=1)
    db.session.add(new_team)

    physicians_string = team.tolist()[list(table_teams).index("Ärzte auf dieser Station")].replace(" ", "").split(',')
    for physician_string in physicians_string:
        new_physician_team = PhysicianStation(institution_id=1, physician_id=[x for x in User.query.all() if x.last_name == physician_string][0].id, station_id=[x for x in Station.query.all() if x == new_team][0].id)
        db.session.add(new_physician_team)


for dutytype in table_duties.values:
    name = dutytype.tolist()[list(table_duties).index("Dienstname")]
    first_day = dutytype.tolist()[list(table_duties).index("Von Tag")]
    last_day = dutytype.tolist()[list(table_duties).index("Bis Tag")]
    day_active = False
    if first_day == "Mo":
        monday = True
        day_active = True
    else:
        monday = False

    if day_active is True or first_day == "Di":
        tuesday = True
        day_active = True
    else:
        tuesday = False
    if last_day == "Di":
        day_active = False

    if day_active is True or first_day == "Mi":
        wednesday = True
        day_active = True
    else:
        wednesday = False
    if last_day == "Mi":
        day_active = False

    if day_active is True or first_day == "Do":
        thursday = True
        day_active = True
    else:
        thursday = False
    if last_day == "Do":
        day_active = False

    if day_active is True or first_day == "Fr":
        friday = True
        day_active = True
    else:
        friday = False
    if last_day == "Fr":
        day_active = False

    if day_active is True or first_day == "Sa":
        saturday = True
        day_active = True
    else:
        saturday = False
    if last_day == "Sa":
        day_active = False

    if day_active is True or first_day == "So":
        sunday = True
        day_active = True
    else:
        sunday = False
    if last_day == "So":
        day_active = False

    time_start = dutytype.tolist()[list(table_duties).index("Uhrzeit von")]
    time_end = dutytype.tolist()[list(table_duties).index("Uhrzeit bis")]

    new_duty_type = DutyType(name=name,
                             institution_id=1,
                             time_start=time_start,
                             time_end=time_end,
                             monday= monday,
                             tuesday= tuesday,
                             wednesday= wednesday,
                             thursday= thursday,
                             friday= friday,
                             saturday= saturday,
                             sunday= sunday,
                             on_holiday=False,
                             before_holiday=False,
                             after_holiday=False,
                             only_on_holiday=False,
                             automatically_create_as_single_duty=True,
                             consecutive_assignment=False,
                             weight_consecutive_assignment=1)

    db.session.add(new_duty_type)

    if dutytype.tolist()[list(table_duties).index("Für Dienst benötigte Qualifikationen")] == dutytype.tolist()[list(table_duties).index("Für Dienst benötigte Qualifikationen")]:
        qualis_strings = dutytype.tolist()[list(table_duties).index("Für Dienst benötigte Qualifikationen")].replace(" ", "").split(',')
        for quali_string in qualis_strings:
            new_dt_quali = DutyQualificationType(institution_id = 1, qualification_id = [x for x in Qualification.query.all() if x.name == quali_string][0].id, duty_type_id=DutyType.query.all()[-1].id)
            db.session.add(new_dt_quali)


for team_duty_type in table_team_duties.values:
    name = team_duty_type.tolist()[list(table_team_duties).index("Stationsdienstname")]
    first_day = team_duty_type.tolist()[list(table_team_duties).index("Von Tag")]
    last_day = team_duty_type.tolist()[list(table_team_duties).index("Bis Tag")]
    day_active = False
    if first_day == "Mo":
        monday = True
        day_active = True
    else:
        monday = False

    if day_active is True or first_day == "Di":
        tuesday = True
        day_active = True
    else:
        tuesday = False
    if last_day == "Di":
        day_active = False

    if day_active is True or first_day == "Mi":
        wednesday = True
        day_active = True
    else:
        wednesday = False
    if last_day == "Mi":
        day_active = False

    if day_active is True or first_day == "Do":
        thursday = True
        day_active = True
    else:
        thursday = False
    if last_day == "Do":
        day_active = False

    if day_active is True or first_day == "Fr":
        friday = True
        day_active = True
    else:
        friday = False
    if last_day == "Fr":
        day_active = False

    if day_active is True or first_day == "Sa":
        saturday = True
        day_active = True
    else:
        saturday = False
    if last_day == "Sa":
        day_active = False

    if day_active is True or first_day == "So":
        sunday = True
        day_active = True
    else:
        sunday = False
    if last_day == "So":
        day_active = False

    time_start = team_duty_type.tolist()[list(table_team_duties).index("Uhrzeit von")]
    time_end = team_duty_type.tolist()[list(table_team_duties).index("Uhrzeit bis")]

    number_of_physicians = team_duty_type.tolist()[list(table_team_duties).index("Anzahl Besetzungen")]
    station_string = team_duty_type.tolist()[list(table_team_duties).index("Station")]
    station_id = [x for x in Station.query.all() if x.name == station_string][0].id

    new_team_type = TeamType(name=name,
             institution_id=1,
             time_start=time_start,
             time_end=time_end,
             monday=monday,
             tuesday=tuesday,
             wednesday=wednesday,
             thursday=thursday,
             friday=friday,
             saturday=saturday,
             sunday=sunday,
             on_holiday=True,
             before_holiday=True,
             after_holiday=True,
             only_on_holiday=False,
             station_id=station_id,
             automatically_create_as_single_team=True,
             number_of_required_physicians=number_of_physicians,
             desired_number_of_physicians=number_of_physicians,
             weight_of_occupation=1)

    db.session.add(new_team_type)

    if team_duty_type.tolist()[list(table_team_duties).index("Für Stationsdienst benötigte Qualifikationen")] == team_duty_type.tolist()[list(table_team_duties).index("Für Stationsdienst benötigte Qualifikationen")]:
        qualis_strings = team_duty_type.tolist()[list(table_team_duties).index("Für Stationsdienst benötigte Qualifikationen")].replace(" ", "").split(',')
        for quali_string in qualis_strings:
            new_team_duty_quali = TeamQualificationType(institution_id=1, team_type_id=TeamType.query.all()[-1].id, qualification_id=[x for x in Qualification.query.all() if x.name == quali_string][0].id, desired_value=True)
            db.session.add(new_team_duty_quali)



for pool in table_pools.values:

    name = pool.tolist()[list(table_pools).index("Poolname")]
    new_pool = Pool(name=name, institution_id=1, fair_distribution=True)
    db.session.add(new_pool)

    physician_string = pool.tolist()[list(table_pools).index("Ärzte in Pool")]
    list_of_physicians_name = physician_string.replace(" ","").split(',')
    pool_id = Pool.query.all()[-1].id
    for physician in list_of_physicians_name:
        physician_id = [p.id for p in list_of_physicians if p.last_name == physician][0]
        new_physician_pool = PhysicianPool(institution_id=1, physician_id=physician_id, pool_id=pool_id)
        db.session.add(new_physician_pool)

    duties_strings = pool.tolist()[list(table_pools).index("Dienste, über die innerhalb dieses Pools fair verteilt werden soll")].replace(" ", "").split(',')
    for duty_string in duties_strings:
        new_pool_dutytype = PoolDutyTypeAssignment(pool_id=pool_id, duty_type_id=[x for x in DutyType.query.all() if x.name.replace(" ", "") == duty_string][0].id)
        db.session.add(new_pool_dutytype)

general_preferences = GeneralPreferencesConfiguration(two_days_on_weekend_preference_allowed=True, two_days_on_weekend_preference_value=100)
db.session.add(general_preferences)

db.session.commit()

