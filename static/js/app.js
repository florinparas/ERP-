// Confirm delete actions
document.addEventListener('DOMContentLoaded', function() {
    // Delete confirmation
    document.querySelectorAll('.btn-delete').forEach(function(btn) {
        btn.addEventListener('click', function(e) {
            if (!confirm('Sigur doriți să ștergeți această înregistrare?')) {
                e.preventDefault();
            }
        });
    });

    // Auto-dismiss alerts after 5 seconds
    document.querySelectorAll('.alert-dismissible').forEach(function(alert) {
        setTimeout(function() {
            var bsAlert = bootstrap.Alert.getOrCreateInstance(alert);
            bsAlert.close();
        }, 5000);
    });

    // Mobile sidebar toggle
    var toggler = document.getElementById('sidebar-toggle');
    if (toggler) {
        toggler.addEventListener('click', function() {
            document.querySelector('.sidebar').classList.toggle('show');
        });
    }
});
