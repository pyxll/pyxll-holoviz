"""
Bridge class that allows PyXLL to display holoviz panels
by serving them in a background thread.
"""
from pyxll import PlotBridgeBase, xl_plot_bridge


_timeout = 15

_template = """
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="refresh" content="0; url='{url}'" />
</head>
<body>
    <p>If you are not redirected automatically, follow this <a href="{url}">link to {url}</a>.</p>
</body>
</html>
"""


@xl_plot_bridge(
    "panel.template.base.BaseTemplate",
    "panel.viewable.Viewable",
    "panel.viewable.Viewer"
)
class HvPanelBridge(PlotBridgeBase):

    def __init__(self, figure):
        PlotBridgeBase.__init__(self, figure)

    def can_export(self, format):
        return format == "html"

    def __get_server(self, future, panel, **kwargs):
        try:
            def started_callback():
                future.set_result(server)

            from panel.io.server import get_server

            # Create the server without (yet) starting it or showing the browser
            server = get_server(panel, start=False, show=False, **kwargs)

            # Start the server (this doesn't start the tornado loop)
            server.start()

            # Add a callback for when the loop's running
            server.io_loop.add_callback(started_callback)

            # And start the tornado loop
            server.io_loop.start()
            
        except Exception as e:
            future.set_exception(e)
            raise

    def export(self, width, height, dpi, format, filename, **kwargs):
        if format != "html":
            raise ValueError("Unable to export as '%s'" % str(format))
        
        # Do the imports here to avoid slowing down Excel starting
        from tornado.ioloop import IOLoop
        import concurrent.futures
        import panel as pn
        import uuid

        # Start the server on a background thread
        loop = IOLoop(make_current=False)

        kwargs.setdefault('server_id', uuid.uuid4().hex)
        future = concurrent.futures.Future()

        thread = pn.io.server.StoppableThread(
            target=self.__get_server,
            io_loop=loop,
            args=(future, self.figure),
            kwargs=kwargs
        )

        thread.daemon = True
        thread.start()

        try:
            # Wait for the server to finish starting up
            server = future.result(timeout=_timeout)

            # And write some html to redirect to our newly started server
            address = server.address or 'localhost'
            url = f"http://{address}:{server.port}{server.prefix}"

            with open(filename, "wt") as fh:
                fh.write(_template.format(url=url))

        except:
            # Stop the thread if there was an error
            thread.stop()
            raise

        # This will be called when the control hosting the page is destroyed
        return thread.stop
