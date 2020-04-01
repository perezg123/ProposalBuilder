$(function(){
    //Listen for a click on any of the dropdown items
    $(".products li").click(function(){
        //Get the value
        var value = $(this).attr("value");
        //Put the retrieved value into the hidden input
        $("input[name='product']").val(value);
    });
});
