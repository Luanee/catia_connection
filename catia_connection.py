# -*- coding: utf-8 -*-
from pycatia import catia
from pycatia.mec_mod_interfaces.part import Part
from pycatia.product_structure_interfaces.product import Product
from psutil import NoSuchProcess, AccessDenied, ZombieProcess, process_iter


class CATIA:
    def __init__(self):
        self.catia = None
        self.catia_process = False
        self.catia_process_list = ["cnext", "catia"]
        self.active_doc = None
        self.active_file = None
        self.count_item = 0
        self.parts = []
        self.products = []
        self.children = {}

        if not self.catia_process:
            self.set_catia_process()

        if self.catia_process:
            self.catia = catia()

        if self.catia and self.catia.documents.count_types(["catpart", "catproduct"]) > 0:
            self.set_active_document()
            self.set_active_file()

    def set_catia_process(self):
        """
        Sets variable if there is any running process that contains on of the given process names.

        """
        process_list = []
        for proc in process_iter():
            try:
                if proc.name().lower().replace(".exe", "") in self.catia_process_list:
                    process_list.append(proc)
            except (OSError, NoSuchProcess, AccessDenied, ZombieProcess):
                pass

        if process_list:
            self.catia_process = True

    def is_catia_active(self):
        """
        Returns true if catia process is running..

        Returns:
            bool: Depending on whether a process was found or not.

        """
        return self.catia_process

    def set_active_document(self):
        """
        Sets the active document.

        """
        self.active_doc = self.catia.active_document

    def get_active_document(self):
        """
        Returns the active document.

        """
        return self.active_doc

    def set_active_file(self):
        """
        Depending on which document exists, the corresponding document class is set.

        """
        if self.active_doc.is_part:
            self.active_file = self.active_doc.part()
        elif self.active_doc.is_product:
            self.active_file = self.active_doc.product()
            self.set_all_children()
        elif self.active_doc.is_drawing:
            self.active_file = self.active_doc.drawing_root()

    def get_active_file(self):
        """
        Returns the active file.

        """
        return self.active_file

    def set_all_children(self, product=None, level=0):
        """Loops through the product and determine if item is part or product. If product was found, the function is executed again.

        Args:
            product (com_object): Contains the part or product class.
            level (int): Provides information about the level of the respective part or product in
                         the structure tree.

        """
        if not product:
            product = self.active_file
            self.add_children(level, product, "CATPart") if product.is_catpart(
            ) else self.add_children(level, product, "CATProduct")

        level += 1

        for i in range(product.count_children()):
            self.count_item += 1
            item = product.get_child(i)
            if item.is_catpart():
                sub_item = Part(item.product)
                self.add_children(level, sub_item, "CATPart")

            elif item.is_catproduct():
                sub_item = Product(item.product)
                self.add_children(level, sub_item, "CATProduct")
                self.set_all_children(sub_item, level)

        self.set_parts()
        self.set_products()

    def add_children(self, level, item, type):
        """Saves every item in dict() 'children' with informations regarding pos, structure level, item, type, mass, volume, wet area, gravity center and inertia.

        Args:
            level (int): Provides information about the level of the respective part or product in
                         the structure tree.
            item (com_object): Contains the part or product class.
            type (str): Contains string with information wheter it is product or part.

        """
        self.children[self.count_item] = {"level": level,
                                          "item": item,
                                          "type": type,
                                          "mass": item.analyze.mass,
                                          "volume": item.analyze.volume,
                                          "wet_area": item.analyze.wet_area}

        gravity_center = item.analyze.get_gravity_center()
        for i in range(3):
            self.children[self.count_item]["G_" + chr(120 + i)] = gravity_center[i]

        inertia = self.get_item_inertia(item)
        inertia_list = ["Ixx", "Ixy", "Ixz", "Iyy", "Iyz", "Izz"]
        for i in range(6):
            self.children[self.count_item][inertia_list[i]] = inertia[i]

    def get_children(self):
        """
        Returns dict() children.

        """
        return self.children

    def get_child(self, index):
        """
        Returns an item that is located at a specific position in the structure tree.

        Args:
            index (int): Position of part or product in structure tree.

        Returns:
            dict: Informations regarding structure level, item, type, mass, volume, wet area, gravity center and inertia.

        """
        try:
            return self.children[index]
        except KeyError:
            return None

    def set_parts(self):
        """
        Places all parts in one list.

        """
        self.parts = [items["item"]
                      for pos, items in self.children.items() if items["type"] == "CATPart"]

    def get_parts(self):
        """
        Returns parts.

        Returns:
            list(com_objects): List of all products

        """
        return self.parts

    def set_products(self):
        """
        Places all products in one list.

        """
        self.products = [items["item"]
                         for pos, items in self.children.items() if items["type"] == "CATProduct"]

    def get_products(self):
        """
        Returns products.

        Returns:
            list(com_objects): List of all products

        """
        return self.products

    def count_product_parts(self):
        """
        Counts the number of all parts.

        Returns:
            int: number of parts.

        """
        if self.parts:
            return len(self.parts)
        return 0

    def count_product_products(self):
        """
        Counts the number of all products.

        Returns:
            int: number of products.

        """
        if self.products:
            return len(self.products)
        return 0

    def find_item(self, keyword):
        """
        Function, which serves to search for a item, wether it is part or product.

        Args:
            keyword (str):  Contains the search keyword.

        Returns:
            com_object: Searched item.

        """
        for child in self.children.values():
            if keyword.lower() in child["item"].name.lower():
                return child["item"]
        return None

    def find_part(self, keyword_part, ignores=[]):
        """
        Function, which serves to search for a product.

        Args:
            keyword_part (str): Contains the search keyword
            ignores (list(str)): Contains strings that schould be ignored in the name of a part

        Returns:
            com_object: contains the searched part or None.

        """
        if not isinstance(keyword_part, list):
            keyword_part = [keyword_part]

        if ignores:
            part = [part for part in self.parts for keyword in keyword_part if keyword.lower(
            ) in part.name.lower() and all(True if part.name.find(ign) < 0 else False for ign in ignores)]
        else:
            part = [part for part in self.parts for keyword in keyword_part if keyword.lower()
                    in part.name.lower()]

        return next(iter(part), None)

    def find_product(self, keyword_product, ignores=[]):
        """
        Function, which serves to search for a product.

        Args:
            keyword_product (str): Contains the search keyword
            ignores (list(str)): Contains strings that schould be ignored in the name of a product

        Returns:
            com_object: contains the searched product or None.

        """
        if not isinstance(keyword_product, list):
            keyword_product = [keyword_product]

        if ignores:
            product = [product for product in self.products for keyword in keyword_product if keyword.lower(
            ) in product.name.lower() and all(True if product.name.find(ign) < 0 else False for ign in ignores)]
        else:
            product = [product for product in self.products for keyword in keyword_product if keyword.lower()
                       in product.name.lower()]

        return next(iter(product), None)

    def get_item_inertia(self, item):
        """
        Function for detecting the essential moments of inertia.

        Args:
            item (com_object): Contains the part or product class.

        Returns:
            tuple(float): Contains Ixx, Ixy, Ixz, Iyy, Iyz, Izz

        """
        inertia = item.analyze.get_inertia()

        return (inertia[0], inertia[1], inertia[2], inertia[4], inertia[5], inertia[8])

    def set_visibility(self, item, hide_param):
        """
        Sets the visibility of specific product

        Args:
            item (com_object): Contains the part or item class.
            hide_param (bool): parameter that indicates whether hidden (False) or visible (True)
                               schould be set

        """
        if item:
            selection = self.active_doc.selection
            selection.add(item)
            visibility = selection.vis_properties
            visibility.set_show(int(hide_param is not True))

    @property
    def version(self):
        """
        Returns:
            str: CATIA version

        """
        return "Version {}".format(self.catia.system_configuration.version)

    @property
    def service_pack(self):
        """
        Returns:
            str: CATIA Service Pack number

        """
        return "Service Pack {}".format(self.catia.system_configuration.service_pack)

    @property
    def release(self):
        """
        Returns:
            str: CATIA Build number

        """
        return "Build Number {}".format(self.catia.system_configuration.release)


if __name__ == "__main__":
    catia_app = CATIA()
    # prints application name or None
    print(catia_app.catia)
    # prints the active document
    print(catia_app.get_active_document())
    # prints the active file
    print(catia_app.get_active_file())
    # find and returns item nr. 1
    item_1 = catia_app.get_child(1)
    # changes the visibility of item_1
    catia_app.set_visibility(item_1, True)
