// ERP System - Client-side JavaScript

document.addEventListener('DOMContentLoaded', function() {
    // Auto-dismiss flash messages after 5 seconds
    document.querySelectorAll('.alert').forEach(function(alert) {
        setTimeout(function() {
            alert.style.opacity = '0';
            alert.style.transition = 'opacity 0.5s';
            setTimeout(function() { alert.remove(); }, 500);
        }, 5000);
    });

    // Confirm delete actions
    document.querySelectorAll('.btn-danger[type="submit"]').forEach(function(btn) {
        btn.addEventListener('click', function(e) {
            if (!confirm('Sunteti sigur ca doriti sa stergeti?')) {
                e.preventDefault();
            }
        });
    });
});

// Add invoice/order item row
function addItemRow() {
    var container = document.getElementById('items-container');
    if (!container) return;
    var index = container.children.length;
    var row = document.createElement('div');
    row.className = 'item-row form-row';
    row.innerHTML =
        '<div class="form-group">' +
            '<input type="text" name="item_description[]" class="form-control" placeholder="Descriere" required>' +
        '</div>' +
        '<div class="form-group">' +
            '<input type="number" name="item_quantity[]" class="form-control" placeholder="Cantitate" step="0.01" value="1" required>' +
        '</div>' +
        '<div class="form-group">' +
            '<input type="number" name="item_price[]" class="form-control" placeholder="Pret" step="0.01" required>' +
        '</div>' +
        '<div class="form-group">' +
            '<input type="number" name="item_vat[]" class="form-control" placeholder="TVA%" value="19">' +
        '</div>' +
        '<div class="form-group" style="align-self:end">' +
            '<button type="button" class="btn btn-danger btn-sm" onclick="this.closest(\'.item-row\').remove()">X</button>' +
        '</div>';
    container.appendChild(row);
}
