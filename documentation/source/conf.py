from sphinx_pyproject import SphinxConfig

# this will inject all the keys from [tool.sphinx-pyproject]
# (and PEP 621 [project] keys like version/description) into globals()
config = SphinxConfig("../pyproject.toml", globalns=globals())
