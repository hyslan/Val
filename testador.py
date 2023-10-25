from sap_connection import connect_to_sap

session = connect_to_sap()


def teste(session):
    def abc():
        # nonlocal session
        session.StartTransaction("mblb")
    abc()


# teste(session)
