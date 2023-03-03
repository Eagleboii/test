Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  // Get an access token for the ezeep blue API
  const accessToken = await getAccessToken();

  // Pair the client with ezeep blue
  const pairingCode = await getPairingCode(accessToken);
  const authUrl = `https://api.ezeep.com/pairing/auth?pairing_code=${pairingCode}`;
  window.open(authUrl);

  // Poll the status of the pairing request until it is complete
  let isPaired = false;
  while (!isPaired) {
    const response = await fetch('https://api.ezeep.com/pairing/status', {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const status = await response.json();
    isPaired = status.status === 'paired';
    await new Promise((resolve) => setTimeout(resolve, 1000)); // Wait for 1 second before polling again
  }

  // Get a list of available printers
  const printers = await getPrinters(accessToken);

  // Add a button to the ribbon that opens the print dialog
  const button = {
    id: 'printButton',
    label: 'Print with ezeep blue',
    iconUrl: 'https://path/to/icon.png',
    onAction: () => showPrintDialog(accessToken, printers),
  };
  const ribbon = Office.context.ui.ribbon;
  ribbon.requestUpdate({
    tabs: [
      {
        id: 'TabHome',
        groups: [
          {
            id: 'GroupPrint',
            label: 'Print',
            controls: [button],
          },
        ],
      },
    ],
  });

  // Function to get an access token for the ezeep blue API
  async function getAccessToken() {
    const response = await fetch('https://login.ezeep.com/oauth/token', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        Authorization: 'Basic IMfGACrsE4xp1MPOZGKFKjQJnoZ6lpsOtmK9ELDm',
      },
      body: 'grant_type=client_credentials',
    });
    const result = await response.json();
    return result.access_token;
  }

  // Function to get a pairing code for ezeep blue
  async function getPairingCode(accessToken) {
    const response = await fetch('https://api.ezeep.com/pairing', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const result = await response.json();
    return result.pairing_code;
  }

  // Function to get a list of available printers
  async function getPrinters(accessToken) {
    const response = await fetch('https://api.ezeep.com/printers', {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const result = await response.json();
    return result.printers;
  }

 // Function to show the print dialog (continued from previous message)
async function showPrintDialog(accessToken, printers) {
  // Create the print dialog UI using React
  const dialogContainer = document.createElement('div');
  ReactDOM.render(
    <PrintDialog accessToken={accessToken} printers={printers} />,
    dialogContainer
  );

  // Define the print dialog UI using React
  function PrintDialog(props) {
    const [selectedPrinterId, setSelectedPrinterId] = useState(null);
    const [numCopies, setNumCopies] = useState(1);
    const [isPrinting, setIsPrinting] = useState(false);

    return (
      <div className="print-dialog">
        <h2>Print with ezeep blue</h2>
        <div className="form-group">
          <label htmlFor="printer-select">Printer:</label>
          <select
            id="printer-select"
            value={selectedPrinterId || ''}
            onChange={(event) => setSelectedPrinterId(event.target.value || null)}
          >
            <option value="">Select a printer</option>
            {props.printers.map((printer) => (
              <option key={printer.id} value={printer.id}>
                {printer.name}
              </option>
            ))}
          </select>
        </div>
        <div className="form-group">
          <label htmlFor="num-copies-input">Number of copies:</label>
          <input
            id="num-copies-input"
            type="number"
            min="1"
            max="100"
            value={numCopies}
            onChange={(event) => setNumCopies(event.target.value)}
          />
        </div>
        <div className="form-group">
          <button
            className="button-primary"
            onClick={() => printDocument(props.accessToken, selectedPrinterId, numCopies, setIsPrinting)}
            disabled={!selectedPrinterId || isPrinting}
          >
            {isPrinting ? 'Printing...' : 'Print'}
          </button>
        </div>
      </div>
    );
  }

  // Function to print the document
  async function printDocument(accessToken, printerId, numCopies, setIsPrinting) {
    setIsPrinting(true);

    const documentUrl = Office.context.document.url;

    const response = await fetch('https://api.ezeep.com/printouts', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        printer_id: printerId,
        document_url: documentUrl,
        num_copies: numCopies,
      }),
    });

    if (response.ok) {
      setIsPrinting(false);
    } else {
      alert('An error occurred while printing the document.');
      setIsPrinting(false);
    }
  }
}
}
