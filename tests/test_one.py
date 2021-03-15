from dash.testing.application_runners import import_app


def test_bbaaa001(dash_duo):
    app = import_app("app")
    dash_duo.start_server(app)

    #dash_duo.wait_for_text_to_equal("/html/body/div/div/div/div[4]/div[1]/div/div/ul/li[1]/a", "Invoer", timeout=10)

    #assert dash_duo.find_element("raillengte").text == "3000"

    #assert dash_duo.get_logs() == [], "Browser console should contain no error"

    return None