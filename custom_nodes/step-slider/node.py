class Step_Controller:
    """
    A simple aspect ratio node

    Class methods
    -------------
    INPUT_TYPES (dict):
        Tell the main program input parameters of nodes.

    Attributes
    ----------
    RETURN_TYPES (`tuple`):
        The type of each element in the output tulple.
    RETURN_NAMES (`tuple`):
        Optional: The name of each output in the output tulple.
    FUNCTION (`str`):
        The name of the entry-point method. For example, if `FUNCTION = "execute"` then it will run Example().execute()
    OUTPUT_NODE ([`bool`]):
        If this node is an output node that outputs a result/image from the graph. The SaveImage node is an example.
        The backend iterates on these output nodes and tries to execute all their parents if their parent graph is properly connected.
        Assumed to be False if not present.
    CATEGORY (`str`):
        The category the node should appear in the UI.
    execute(s) -> tuple || None:
        The entry point method. The name of this method must be the same as the value of property `FUNCTION`.
        For example, if `FUNCTION = "execute"` then this method's name must be `execute`, if `FUNCTION = "foo"` then it must be `foo`.
    """

    def __init__(self):
        pass

    @classmethod
    def INPUT_TYPES(s):
        """
        Return a dictionary which contains config for all input fields.
        """
        return {
            "required": {
                "diffusion_step": ("INT", {
                    "default": 1, 
                    "min": 1, 
                    "max": 20, 
                    "step": 1, 
                    "display": "slider"  # Display as a slider
                }),
            }
        }

    RETURN_TYPES = ("diffusion_step",)
    RETURN_NAMES = ("Step",)

    FUNCTION = "execute"

    CATEGORY = "utils"

    def execute(self, diffusion_step):
        # Your processing logic
        return (diffusion_step,)

# A dictionary that contains all nodes you want to export with their names
# NOTE: names should be globally unique
NODE_CLASS_MAPPINGS = {"Step_Controller": Step_Controller}

# A dictionary that contains the friendly/humanly readable titles for the nodes
NODE_DISPLAY_NAME_MAPPINGS = {"Step_Controller": "Diffusion Controller"}
