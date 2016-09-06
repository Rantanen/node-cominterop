var native = require( 'bindings' )( 'addon' );

module.exports.load = function( path ) {

    // Load the native library.
    var lib = native.load( path );

    // Enhance the types.
    Object.keys( lib ).forEach( function( obj ) {

        // Go through all prototype methods.
        var cls = lib[ obj ];

        // Gather all gettable/settable properties.
        var properties = {};
        Object.keys( cls.prototype ).forEach( function( funcName ) {

            var func = cls.prototype[ funcName ];

            // Only interested in getter/setter properties.
            var propName = func.getter || func.setter;
            if( !propName )
                return;

            // Ensure we got the property tracked.
            properties[ propName ] = properties[ propName ] || {};
            
            // Check whether the current function is getter or setter.
            if( func.getter ) {

                // Getter. Define the getter alias.
                properties[ propName ].get = function() {
                    return this[ funcName ].apply( this, arguments );
                };

            } else if( func.setter ) {
                
                // Setter. Define the setter alias.
                properties[ propName ].set = function() {
                    return this[ funcName ].apply( this, arguments );
                };
            }
        } );

        // Define all the proeprties on the object prototype.
        Object.keys( properties ).forEach( function( propName ) {
            Object.defineProperty( cls.prototype, propName, properties[ propName ] );
        } );
    } );

    return lib;
};

