import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './ValidationFormWebPart.module.scss';

export default class ValidationFormWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    this.domElement.innerHTML = `
<div class="${styles['form-container']}">
    <h1>Form Validation</h1>
    <div class="${styles['Formcontrols']}">

        <!-- BID DETAILS Section -->
        <div class="${styles['BidDetails']}">
            <button type="button" class="toggleBtn" data-target="bidForm">BID DETAILS</button>
            <form id="bidForm" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="fname1">How long is the response time?</label>
                    <input type="text" id="fname1" name="fname1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="lname1">Do you have the capacity to submit a compliant response on time?</label>
                    <input type="text" id="lname1" name="lname1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="Que3">Does the bid appear to be written for another organisation?</label>
                    <input type="text" id="Que3" name="Que3">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="Que4">What is the evaluation criteria?</label>
                    <input type="text" id="Que4" name="Que4">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="Que5">Do we meet the specific goals pre-qualifying criteria?</label>
                    <input type="text" id="Que5" name="Que5">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="Que6">What is the bid validity period?</label>
                    <input type="text" id="Que6" name="Que6">
                </div>
            </form>
        </div>

        <!-- CLIENTS Section -->
        <div class="${styles['Clients']}">
            <button type="button" class="toggleBtn" data-target="bidClient">CLIENTS</button>
            <form id="bidClient" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue1">What are the main requirements and objectives in the RFP?</label>
                    <input type="text" id="ClientQue1" name="ClientQue1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue2">How well do we understand their requirements?</label>
                    <input type="text" id="ClientQue2" name="ClientQue2">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue3">Do we have a relationship with the client?</label>
                    <input type="text" id="ClientQue3" name="ClientQue3">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue4">Who's our point of contact and what's the contact information?</label>
                    <input type="text" id="ClientQue4" name="ClientQue4">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue5">Do we have experience working with this client?</label>
                    <input type="text" id="ClientQue5" name="ClientQue5">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue6">If so, list projects</label>
                    <input type="text" id="ClientQue6" name="ClientQue6">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue7">Have we performed a similar project before (value and size)?</label>
                    <input type="text" id="ClientQue7" name="ClientQue7">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue8">Are we a good fit for this client?</label>
                    <input type="text" id="ClientQue8" name="ClientQue8">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue9">What will we benefit from this opportunity (experience in a new industry, developing a new client relationship, etc.)?</label>
                    <input type="text" id="ClientQue9" name="ClientQue9">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="ClientQue10">Is there any adverse media/ conflicts about the client which could impact us?</label>
                    <input type="text" id="ClientQue10" name="ClientQue10">
                </div>
            </form>
        </div>

        <!-- COMPETITION Section -->
        <div class="${styles['Clients']}">
            <button type="button" class="toggleBtn" data-target="competition">THE COMPETITION</button>
            <form id="competition" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="CompetitionQue1_1">Who is the current incumbent?</label>
                    <input type="text" id="CompetitionQue1_1" name="CompetitionQue1_1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="CompetitionQue2_1">Why is the client seeking a new supplier?</label>
                    <input type="text" id="CompetitionQue2_1" name="CompetitionQue2_1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="CompetitionQue3_1">List all known competitors / bidders.</label>
                    <input type="text" id="CompetitionQue3_1" name="CompetitionQue3_1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="CompetitionQue4_1">Has the competition worked with the client in the past?</label>
                    <input type="text" id="CompetitionQue4_1" name="CompetitionQue4_1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="CompetitionQue5_1">Compared to us, do one or more competitors provide a superior service or solution?</label>
                    <input type="text" id="CompetitionQue5_1" name="CompetitionQue5_1">
                </div>
            </form>
        </div>

        <!-- OUR ORGANISATION -->
        <div class="${styles['Clients']}">
            <button type="button" class="toggleBtn" data-target="organisation">OUR ORGANISATION</button>
            <form id="organisation" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="orgQue1">Who is the current incumbent?</label>
                    <input type="text" id="orgQue1" name="orgQue1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="orgQue2">Do we have the resources available to respond?</label>
                    <input type="text" id="orgQue2" name="orgQue2">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="orgQue3">Do we have the resources available to implement the solution?</label>
                    <input type="text" id="orgQue3" name="orgQue3">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="orgQue4">If not, can we partner with someone who does?</label>
                    <input type="text" id="orgQue4" name="orgQue4">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="orgQue5">Are we currently marketing to this client?</label>
                    <input type="text" id="orgQue5" name="orgQue5">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="orgQue6">Have we influenced the RFP?</label>
                    <input type="text" id="orgQue6" name="orgQue6">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="orgQue7">What is the probability of winning this bid?</label>
                    <input type="text" id="orgQue7" name="orgQue7">
                </div>
            </form>
        </div>

        <!-- OUR SOLUTION -->
        <div class="${styles['Clients']}">
            <button type="button" class="toggleBtn" data-target="solution">OUR SOLUTION</button>
            <form id="solution" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="solQue1">Does our solution fit the client's requirements and objectives?</label>
                    <input type="text" id="solQue1" name="solQue1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="solQue2">Is the solution within our area(s) of expertise?</label>
                    <input type="text" id="solQue2" name="solQue2">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="solQue3">Is this type of work desirable for us?</label>
                    <input type="text" id="solQue3" name="solQue3">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="solQue4">Is the location of performance a good fit for us?</label>
                    <input type="text" id="solQue4" name="solQue4">
                </div>
            </form>
        </div>

        <!-- OUR STRATEGIC OBJECTIVES -->
        <div class="${styles['Clients']}">
            <button type="button" class="toggleBtn" data-target="stratObjective">OUR STRATEGIC OBJECTIVES</button>
            <form id="stratObjective" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="stratOb1">Does our solution fit the client's requirements and objectives?</label>
                    <input type="text" id="stratOb1" name="stratOb1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="stratOb2">Is the solution within our area(s) of expertise?</label>
                    <input type="text" id="stratOb2" name="stratOb2">
                </div>
            </form>
        </div>

        <!-- THE RISKS Section -->
        <div class="${styles['Clients']}">
            <button type="button" class="toggleBtn" data-target="risks">THE RISKS</button>
            <form id="risks" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="risk1">What risks are associated with winning?</label>
                    <input type="text" id="risk1" name="risk1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="risk2">Has the client assigned a budget for the work?</label>
                    <input type="text" id="risk2" name="risk2">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="risk3">What, if any, are the penalties for nonperformance?</label>
                    <input type="text" id="risk3" name="risk3">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="risk4">Are the known risks acceptable?</label>
                    <input type="text" id="risk4" name="risk4">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="risk5">If the risks are not acceptable, can they be mitigated or shifted?</label>
                    <input type="text" id="risk5" name="risk5">
                </div>
            </form>
        </div>

        <!-- OUR FINANCES Section -->
        <div class="${styles['Clients']}">
            <button type="button" class="toggleBtn" data-target="finances">OUR FINANCES</button>
            <form id="finances" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="finance1">Is this a profitable project?</label>
                    <input type="text" id="finance1" name="finance1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="finance2">Do we have the budget to respond to this opportunity?</label>
                    <input type="text" id="finance2" name="finance2">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="finance3">What is the estimated cost of providing the solution?</label>
                    <input type="text" id="finance3" name="finance3">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="finance4">What do we know about the client's payment practices?</label>
                    <input type="text" id="finance4" name="finance4">
                </div>
            </form>
        </div>

        <!-- OUR PARTNERS Section -->
        <div class="${styles['Clients']}">
            <button type="button" class="toggleBtn" data-target="partners">OUR PARTNERS</button>
            <form id="partners" style="display: none;">
                <div class="${styles['form-inputs']}">
                    <label for="partner1">Do we need a partner?</label>
                    <input type="text" id="partner1" name="partner1">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="partner2">Have we identified the partner?</label>
                    <input type="text" id="partner2" name="partner2">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="partner3">Have we worked with this partner in the past?</label>
                    <input type="text" id="partner3" name="partner3">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="partner4">How well do we trust this partner?</label>
                    <input type="text" id="partner4" name="partner4">
                </div>
                <div class="${styles['form-inputs']}">
                    <label for="partner5">Is the partner available to assist with response and performance?</label>
                    <input type="text" id="partner5" name="partner5">
                </div>
            </form>
        </div>
<div class="${styles['Buttondiv']}">
        <button value="Submit" type="submit" id="submitBtn">Submit</button>
        </div>
    </div>
</div>


    `;

    this._attachEventListeners();
  }

  private _attachEventListeners(): void {
    // Toggle form visibility
    const toggleButtons = this.domElement.querySelectorAll('.toggleBtn');
    toggleButtons.forEach(button => {
      button.addEventListener('click', (event: Event) => {
        const target = (event.target as HTMLElement).getAttribute('data-target');
        if (target) {
          const form = this.domElement.querySelector(`#${target}`) as HTMLElement;
          if (form) {
            form.style.display = form.style.display === 'block' ? 'none' : 'block';
          }
        }
      });
    });

    // Submit button listener
    const submitButton = this.domElement.querySelector('#submitBtn') as HTMLButtonElement;
    submitButton.addEventListener('click', (event: Event) => {
      event.preventDefault();
      this._submitForm();
    });
  }

  private _submitForm(): void {
    // Collect form values of BID DETAILS
    const fname1 = (document.getElementById('fname1') as HTMLInputElement).value;
    const lname1 = (document.getElementById('lname1') as HTMLInputElement).value;
    const Que3 = (document.getElementById('Que3') as HTMLInputElement).value;
    const Que4 = (document.getElementById('Que4') as HTMLInputElement).value;
    const Que5 = (document.getElementById('Que5') as HTMLInputElement).value;
    const Que6 = (document.getElementById('Que6') as HTMLInputElement).value;

    // Collect form values of CLIENTS Section
    const ClientQue1 = (document.getElementById('ClientQue1') as HTMLInputElement).value;
    const ClientQue2 = (document.getElementById('ClientQue2') as HTMLInputElement).value;
    const ClientQue3 = (document.getElementById('ClientQue3') as HTMLInputElement).value;
    const ClientQue4 = (document.getElementById('ClientQue4') as HTMLInputElement).value;
    const ClientQue5 = (document.getElementById('ClientQue5') as HTMLInputElement).value;
    const ClientQue6 = (document.getElementById('ClientQue6') as HTMLInputElement).value;
    const ClientQue7 = (document.getElementById('ClientQue7') as HTMLInputElement).value;
    const ClientQue8 = (document.getElementById('ClientQue8') as HTMLInputElement).value;
    const ClientQue9 = (document.getElementById('ClientQue9') as HTMLInputElement).value;
    const ClientQue10 = (document.getElementById('ClientQue10') as HTMLInputElement).value;

     // Collect form values of COMPETITION Section 
     const CompetitionQue1_1 = (document.getElementById('CompetitionQue1_1') as HTMLInputElement).value;
     const CompetitionQue2_1 = (document.getElementById('CompetitionQue2_1') as HTMLInputElement).value;
     const CompetitionQue3_1 = (document.getElementById('CompetitionQue3_1') as HTMLInputElement).value;
     const CompetitionQue4_1 = (document.getElementById('CompetitionQue4_1') as HTMLInputElement).value;
     const CompetitionQue5_1 = (document.getElementById('CompetitionQue5_1') as HTMLInputElement).value;

 // Collect form values of OUR ORGANISATION 
 const orgQue1 = (document.getElementById('orgQue1') as HTMLInputElement).value;
 const orgQue2 = (document.getElementById('orgQue2') as HTMLInputElement).value;
 const orgQue3 = (document.getElementById('orgQue3') as HTMLInputElement).value;
 const orgQue4 = (document.getElementById('orgQue4') as HTMLInputElement).value;
 const orgQue5 = (document.getElementById('orgQue5') as HTMLInputElement).value;
 const orgQue6 = (document.getElementById('orgQue6') as HTMLInputElement).value;
 const orgQue7 = (document.getElementById('orgQue7') as HTMLInputElement).value;

  // Collect form values of OUR SOLUTION
  const solQue1 = (document.getElementById('solQue1') as HTMLInputElement).value;
  const solQue2 = (document.getElementById('solQue2') as HTMLInputElement).value;
  const solQue3 = (document.getElementById('solQue3') as HTMLInputElement).value;
  const solQue4 = (document.getElementById('solQue4') as HTMLInputElement).value;

  //Collect form values of OUR STRATEGIC OBJECTIVES
  const stratOb1 = (document.getElementById('stratOb1') as HTMLInputElement).value;
  const stratOb2 = (document.getElementById('stratOb2') as HTMLInputElement).value;

  //Collect form values of THE RISKS Section
  const risk1 = (document.getElementById('risk1') as HTMLInputElement).value;
  const risk2 = (document.getElementById('risk2') as HTMLInputElement).value;
  const risk3 = (document.getElementById('risk3') as HTMLInputElement).value;
  const risk4 = (document.getElementById('risk4') as HTMLInputElement).value;
  const risk5 = (document.getElementById('risk5') as HTMLInputElement).value;

  //Collect form values of OUR FINANCE
  const finance1 = (document.getElementById('finance1') as HTMLInputElement).value;
  const finance2 = (document.getElementById('finance2') as HTMLInputElement).value;
  const finance3 = (document.getElementById('finance3') as HTMLInputElement).value;
  const finance4 = (document.getElementById('finance4') as HTMLInputElement).value;


    //Collect form values of OUR PARTNERS
    const partner1 = (document.getElementById('partner1') as HTMLInputElement).value;
    const partner2 = (document.getElementById('partner2') as HTMLInputElement).value;
    const partner3 = (document.getElementById('partner3') as HTMLInputElement).value;
    const partner4= (document.getElementById('partner4') as HTMLInputElement).value;
    const partner5= (document.getElementById('partner5') as HTMLInputElement).value;
  
    if (!fname1 && !lname1) {
      alert('Please fill in at least one field.');
      return;
    }

    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const listName = "TenderQuestionnaires"; // Replace with your SharePoint list name

    const data = {
      "Howlongistheresponsetime_x003f_": fname1, // Replace with actual column names
      "Doyoucapacitytosubmitacompliantr":lname1,
      "Doesthebidappeartobewrittenforan":Que3,
      "Whatistheevaluationcriteria_x003":Que4,
      "Dowemeetthespecicgoalspre_x002d_":Que5,
      "Whatisthebidvalidityperiod_x003f":Que6,
      "Whatarethemainrequirementsandobj":ClientQue1,
      "Howwelldoweunderstandtheirrequir":ClientQue2,
      "Dowehavearelationshipwiththeclie":ClientQue3,
      "Whosourpointofcontactandwhatsthe":ClientQue4,
      "Dowehaveexperienceworkingwiththi":ClientQue5,
      "Ifso_x002c_listprojects_x002e_":ClientQue6,
      "Haveweperformedasimilarprojectbe":ClientQue7,
      "Areweagoodfitforthisclient_x003f":ClientQue8,
      "Whatwillwebenefitfromthisopportu":ClientQue9,
      "Isthereanyadversemedia_x002f_con":ClientQue10,
      "Whoisthecurrentincumbent_x003f_":CompetitionQue1_1,
      "Whyistheclientseekinganewsupplie":CompetitionQue2_1,
      "Listallknowncompetitors_x002f_bi":CompetitionQue3_1,
      "Hasthecompetitionworkedwiththecl":CompetitionQue4_1,
      "Comparedtous_x002c_dooneormoreco":CompetitionQue5_1,
      "Wouldwerespondasprimeorsub_x003f":orgQue1,
      "Dowehavetheresourcesavailabletor":orgQue2,
      "Dowehavetheresourcesavailabletoi":orgQue3,
      "Ifnot_x002c_canwepartnerwithsome":orgQue4,
      "Arewecurrentlymarketingtothiscli":orgQue5,
      "HaveweinfluencedtheRFP_x003f_":orgQue6,
      "Whatistheprobabilityofwinningthi":orgQue7,
      "Doesoursolutionfittheclientsrequ":solQue1,
      "Isthesolutionwithinourarea_x0028":solQue2,
      "Isthistypeofworkdesireableforus_":solQue3,
      "Isthelocationofperformanceagoodf":solQue4,
      "Doesthisopportunityalignwithours":stratOb1,
      "Wouldwinningprovideaccesstoanewm":stratOb2,
      "Whatrisksareassociatedwithwinnin":risk1,
      "Hastheclientassignedabudgetforth":risk2,
      "What_x002c_ifany_x002c_arethepen":risk3,
      "Aretheknownrisksacceptable_x003f":risk4,
      "Iftherisksarenotacceptable_x002c":risk5,
      "Isthisaprofitableproject_x003f_":finance1,
      "Dowehavethebudgettorespondtothis":finance2,
      "Whatistheestimatedcostofprovidin":finance3,
      "Whatdoweknowabouttheclientspayme":finance4,
      "Doweneedapartner_x003f_":partner1,
      "Haveweidentifiedthepartner_x003f":partner2,
      "Haveweworkedwiththispartnerinthe":partner3,
      "Howwelldowetrustthispartner_x003":partner4,
      "Isthepartneravailabletoassistwit":partner5,
    };

    // Get form digest for authentication
    fetch(`${siteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: { 'Accept': 'application/json;odata=verbose' }
    })
      .then(response => response.json())
      .then(contextInfo => {
        const requestDigest = contextInfo.d.GetContextWebInformation.FormDigestValue;

        // Submit data to SharePoint list
        return fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
          method: 'POST',
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest
          },
          body: JSON.stringify({
            "__metadata": { "type": `SP.Data.${listName}ListItem` }, // Replace with actual list item type
            ...data
          })
        });
      })
      .then(response => {
        if (!response.ok) {
          return response.json().then(error => {
            throw new Error(error.error.message.value);
          });
        }
        return response.json();
      })
      .then(result => {
        alert('Data successfully submitted!');
      })
      .catch(error => {
        console.error('Error submitting data:', error);
        alert('Error submitting data: ' + error.message);
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
