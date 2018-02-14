<?php
/**
 * The base configuration for WordPress
 *
 * The wp-config.php creation script uses this file during the
 * installation. You don't have to use the web site, you can
 * copy this file to "wp-config.php" and fill in the values.
 *
 * This file contains the following configurations:
 *
 * * MySQL settings
 * * Secret keys
 * * Database table prefix
 * * ABSPATH
 *
 * @link https://codex.wordpress.org/Editing_wp-config.php
 *
 * @package WordPress
 */

// ** MySQL settings - You can get this info from your web host ** //
/** The name of the database for WordPress */
define('DB_NAME', 'mybox2hj_wp1');

/** MySQL database username */
define('DB_USER', 'mybox2hj_wp1');

/** MySQL database password */
define('DB_PASSWORD', 'ggfahTokG');

/** MySQL hostname */
define('DB_HOST', 'localhost');

/** Database Charset to use in creating database tables. */
define('DB_CHARSET', 'utf8mb4');

/** The Database Collate type. Don't change this if in doubt. */
define('DB_COLLATE', '');

/**#@+
 * Authentication Unique Keys and Salts.
 *
 * Change these to different unique phrases!
 * You can generate these using the {@link https://api.wordpress.org/secret-key/1.1/salt/ WordPress.org secret-key service}
 * You can change these at any point in time to invalidate all existing cookies. This will force all users to have to log in again.
 *
 * @since 2.6.0
 */
define('AUTH_KEY',         'Wg E<Ci,`x=gbW Y%ST6n2PWi=9x6#]K`3ki=XpkiNcgarz^h@X4`:>Bhp|xm(9~');
define('SECURE_AUTH_KEY',  'Q9p(zg H<5 xEz^_%WQTO0G.?xd/B{*o},Z<Ev&K&(%ql-6)ZCP^ BycS7V&4<ZS');
define('LOGGED_IN_KEY',    '|Ab[6=k|^ta+J-cdEY{MK2vG;_5?^gDR<zl1*hjz7UOC>Xt_h:~zbJ-N4Q-?:u|/');
define('NONCE_KEY',        'O5<EbFYJ4<Oy|O?m`Z[c.6sZ/W+/V$z%mpbmQwBLpuB}[;:% NR1r]$Y^mQu,2A8');
define('AUTH_SALT',        'h$Jd@z4Lt>,8-A+?8wL*|z5!a[g$PM$[B4$,=9CFO01?zFuqa)?lhH!l[3{4K4ba');
define('SECURE_AUTH_SALT', 'RAuxdK]8 2{Fr+zQ7N9Q}qN`EgggX:!A1gVDbrGU[zQu+h;Cmx2s4{CKN&ZHPlcC');
define('LOGGED_IN_SALT',   '3^`zJ8M`bIX&`jn]n&E(0gn/a_jj@^Dt[~$G,udH%Gkr#$OLVkODl9.>Dvf3Mq$c');
define('NONCE_SALT',       'fsI9o67Xu>p/sL[-sQ2cjlw9BJ-<^g*|uVIU^0;cV3Z`RQ5;L>Vu~Dg]alCavwYI');

/**#@-*/

/**
 * WordPress Database Table prefix.
 *
 * You can have multiple installations in one database if you give each
 * a unique prefix. Only numbers, letters, and underscores please!
 */
$table_prefix  = 'wp_';

/**
 * For developers: WordPress debugging mode.
 *
 * Change this to true to enable the display of notices during development.
 * It is strongly recommended that plugin and theme developers use WP_DEBUG
 * in their development environments.
 *
 * For information on other constants that can be used for debugging,
 * visit the Codex.
 *
 * @link https://codex.wordpress.org/Debugging_in_WordPress
 */
define('WP_DEBUG', false);

/* That's all, stop editing! Happy blogging. */

/** Absolute path to the WordPress directory. */
if ( !defined('ABSPATH') )
	define('ABSPATH', dirname(__FILE__) . '/');

/** Sets up WordPress vars and included files. */
require_once(ABSPATH . 'wp-settings.php');