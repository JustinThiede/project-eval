$(function(){
    $('.add').on('click', addFields);
    $(document).on('click', '.remove', removeField);

    /**
     * Add form fields
     */
    function addFields() {
        var $lastRecord = $('.recordset:last');

        $lastRecord.clone().insertAfter($lastRecord).find('input').val('');
    }

    /**
     * Remove form fields
     */
    function removeField() {
        // Atleast one dataset needs to be listed
        if ($('.recordset').length > 1) {
            $(this).parents('.recordset').remove();
        } else {
            alert('Atleast one dataset needs to be created.');
        }
    }
});
