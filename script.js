document.getElementById('emailForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    // Get form elements
    const excelFile = document.getElementById('excel_file').files[0];
    const senderEmail = document.getElementById('sender_email').value;
    const senderPassword = document.getElementById('sender_password').value;
    const subject = document.getElementById('subject').value;
    const content = document.getElementById('content').value;
    
    // Create FormData object
    const formData = new FormData();
    formData.append('excel_file', excelFile);
    formData.append('sender_email', senderEmail);
    formData.append('sender_password', senderPassword);
    formData.append('subject', subject);
    formData.append('content', content);
    
    // Show progress section
    document.getElementById('progress').style.display = 'block';
    document.getElementById('results').style.display = 'none';
    document.getElementById('sendButton').disabled = true;
    
    try {
        const response = await fetch('/send_emails', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            // Update progress bar to 100%
            document.querySelector('.progress-fill').style.width = '100%';
            
            // Display results
            const resultsList = document.getElementById('resultsList');
            resultsList.innerHTML = '';
            
            data.results.forEach(result => {
                const resultItem = document.createElement('div');
                resultItem.className = `result-item ${result.status.toLowerCase()}`;
                resultItem.textContent = `${result.email}: ${result.status} - ${result.message}`;
                resultsList.appendChild(resultItem);
            });
            
            // Update status
            document.getElementById('status').textContent = 
                `Completed: ${data.successful} out of ${data.total} emails sent successfully`;
            
            // Show results section
            document.getElementById('results').style.display = 'block';
        } else {
            throw new Error(data.error || 'Failed to send emails');
        }
    } catch (error) {
        document.getElementById('status').textContent = `Error: ${error.message}`;
        document.querySelector('.progress-fill').style.backgroundColor = '#dc3545';
    } finally {
        document.getElementById('sendButton').disabled = false;
    }
});
