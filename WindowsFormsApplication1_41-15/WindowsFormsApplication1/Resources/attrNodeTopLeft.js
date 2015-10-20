attrNodeTopLeft = function() {
    $('body *:not(script, noscript)').each(function() {
        var o = $(this);
        o.attr('top', o.offset().top)
            .attr('left', o.offset().left)
            .attr('height', o.innerHeight())
            .attr('width', o.innerWidth())
    });
}