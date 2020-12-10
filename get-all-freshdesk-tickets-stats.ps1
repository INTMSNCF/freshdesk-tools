[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$subdomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your subdomain:`nex.: (intm)", 'Server info').ToLower()
$token = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your freshdesk token:", 'Auth')

$password = "X"
$startDate = (get-date).AddDays(-365)

$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $token,$password)))

try
{
    $response = Invoke-RestMethod -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -Uri "https://$($subdomain).freshdesk.com/api/v2/tickets?updated_since=$($startDate.ToString("yyyy-MM-dd"))&include=stats"
}
catch
{
    write-host -ForegroundColor DarkRed $_.Exception.Message
    read-host “Press ENTER to end...”
    exit 1
}

$database = [System.Collections.ArrayList]::new()

$response | % {
    $item = [pscustomobject]@{
        cc_emails = [system.String]::Join(",", $_.cc_emails)
        fwd_emails = [system.String]::Join(",", $_.fwd_emails)
        reply_cc_emails = [system.String]::Join(",", $_.reply_cc_emails)
        ticket_cc_emails = [system.String]::Join(",", $_.ticket_cc_emails)
        fr_escalated = $_.fr_escalated
        spam = $_.spam
        email_config_id = $_.email_config_id
        group_id = $_.group_id
        priority = $_.priority
        requester_id = $_.requester_id
        responder_id = $_.responder_id
        source = $_.source
        company_id = $_.company_id
        status = $_.status
        subject = $_.subject
        association_type = $_.association_type
        to_emails = $_.to_emails
        product_id = $_.product_id
        id = $_.id
        type = $_.type
        due_by = $_.due_by
        fr_due_by = $_.fr_due_by
        is_escalated = $_.is_escalated
        cf_version = $_.cf_version
        cf_service_demandeur328101 = $_.cf_service_demandeur328101
        cf_service_demandeur = $_.cf_service_demandeur
        cf_logicielle = $_.cf_logicielle
        cf_si_autre_version = $_.cf_si_autre_version
        cf_criticit = $_.cf_criticit
        cf_fsm_contact_name = $_.cf_fsm_contact_name
        cf_fsm_phone_number = $_.cf_fsm_phone_number
        cf_fsm_service_location = $_.cf_fsm_service_location
        cf_fsm_appointment_start_time = $_.cf_fsm_appointment_start_time
        cf_fsm_appointment_end_time = $_.cf_fsm_appointment_end_time
        created_at = $_.created_at
        updated_at = $_.updated_at
        associated_tickets_count = $_.associated_tickets_count
        tags = [system.String]::Join(",", $_.tags)
        agent_responded_at = $_.agent_responded_at
        requester_responded_at = $_.requester_responded_at
        first_responded_at = $_.first_responded_at
        status_updated_at = $_.status_updated_at
        reopened_at = $_.reopened_at
        resolved_at = $_.resolved_at
        closed_at = $_.closed_at
        pending_since = $_.pending_since
        nr_due_by = $_.nr_due_by
        nr_escalated = $_.nr_escalated
    }
    [void]$database.Add($item)
}

$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog.Title = 'Destino'
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.filter = 'Comma-separated values (*.csv)| *.csv'
$OpenFileDialog.ShowDialog() | Out-Null
$exportCsv = $OpenFileDialog.filename

$database | Export-Csv -Path $exportCsv -Delimiter ';'
