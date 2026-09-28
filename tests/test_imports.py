from ms_graph_toolbox.graph_email import send_email
from ms_graph_toolbox.graph_sharepoint import graph_sharepoint
from ms_graph_toolbox.graph_users import get_users
from ms_graph_toolbox.ms_graph_toolbox import ms_graph_toolbox


def test_public_modules_import():
    assert callable(send_email)
    assert callable(get_users)
    assert callable(graph_sharepoint)
    assert callable(ms_graph_toolbox)