import unittest
from unittest.mock import Mock, patch

from ttpy.tournamentroutines import _create_client, _get_credentials


class TournamentRegistrationHelperTests(unittest.TestCase):
    @patch('ttpy.tournamentroutines.helpers.serialize_object')
    @patch('ttpy.tournamentroutines.zeep.Client')
    def test_create_client_initializes_soap_client(self, client_factory, serialize):
        client = client_factory.return_value
        tournaments = client.service.GetTournaments.return_value
        serialize.return_value = {'TournamentEntries': []}

        result_client, result_data = _create_client()

        client_factory.assert_called_once_with(
            wsdl='http://api.vttl.be/0.7/?wsdl')
        client.service.GetTournaments.assert_called_once_with()
        serialize.assert_called_once_with(tournaments)
        self.assertIs(result_client, client)
        self.assertEqual(result_data, {'TournamentEntries': []})

    @patch.dict(
        'os.environ',
        {'ACCOUNT_TT': 'account', 'PASWOORD_TT': 'password'},
        clear=True,
    )
    @patch('builtins.input', new_callable=Mock)
    def test_get_credentials_reads_environment(self, input_mock):
        self.assertEqual(_get_credentials(), ('account', 'password'))
        input_mock.assert_not_called()


if __name__ == '__main__':
    unittest.main()
